const express = require('express');
const multer = require('multer');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const csv = require('csv-parser');
const { v4: uuidv4 } = require('uuid');
const XLSX = require('xlsx');
const PDFDocument = require('pdfkit');
const natural = require('natural');
const stats = require('simple-statistics');

const app = express();
const DEFAULT_PORT = 3000;

app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static('public'));

const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        const uploadDir = './uploads';
        if (!fs.existsSync(uploadDir)) {
            fs.mkdirSync(uploadDir);
        }
        cb(null, uploadDir);
    },
    filename: function (req, file, cb) {
        cb(null, Date.now() + '-' + file.originalname);
    }
});

const upload = multer({
    storage: storage,
    fileFilter: function (req, file, cb) {
        if (file.mimetype === 'text/csv' || file.originalname.endsWith('.csv') || file.originalname.endsWith('.xlsx')) {
            cb(null, true);
        } else {
            cb(new Error('Only CSV and Excel files are allowed'));
        }
    },
    limits: {
        fileSize: 100 * 1024 * 1024
    }
});

let uploadedData = [];
let currentHeaders = [];
let dataId = null;
let columnTypes = {};
let dataSummary = {};
let cleanedData = [];
let dataHistory = [];
let savedSearches = [];

// ============ ROUTES ============

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Upload File
app.post('/api/upload', upload.single('file'), (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'No file uploaded' });
        }

        const filePath = req.file.path;
        const ext = path.extname(req.file.originalname).toLowerCase();
        let results = [];

        if (ext === '.csv') {
            fs.createReadStream(filePath)
                .pipe(csv())
                .on('data', (data) => results.push(data))
                .on('end', () => {
                    fs.unlinkSync(filePath);
                    processUploadedData(results, res);
                })
                .on('error', (err) => {
                    fs.unlinkSync(filePath);
                    res.status(500).json({ error: 'Error processing CSV: ' + err.message });
                });
        } else if (ext === '.xlsx') {
            const workbook = XLSX.readFile(filePath);
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            results = XLSX.utils.sheet_to_json(worksheet);
            fs.unlinkSync(filePath);
            processUploadedData(results, res);
        } else {
            fs.unlinkSync(filePath);
            res.status(400).json({ error: 'Unsupported file format' });
        }
    } catch (error) {
        res.status(500).json({ error: 'Server error: ' + error.message });
    }
});

function processUploadedData(results, res) {
    if (results.length === 0) {
        return res.status(400).json({ error: 'File is empty or invalid' });
    }

    dataId = uuidv4();
    uploadedData = results.map((row, index) => ({
        id: uuidv4(),
        _index: index,
        ...row
    }));
    currentHeaders = Object.keys(results[0]);
    cleanedData = [...uploadedData];

    columnTypes = detectColumnTypes(uploadedData);
    dataSummary = generateDataSummary(uploadedData, columnTypes);
    const stats = calculateStats(uploadedData);

    dataHistory.push({
        id: uuidv4(),
        timestamp: new Date().toISOString(),
        action: 'upload',
        data: uploadedData,
        headers: currentHeaders
    });

    res.json({
        success: true,
        message: `Successfully uploaded ${results.length} rows`,
        data: uploadedData,
        headers: currentHeaders,
        stats: stats,
        columnTypes: columnTypes,
        summary: dataSummary,
        dataId: dataId
    });
}

// Get Data with search, filter, pagination
app.get('/api/data', (req, res) => {
    try {
        const search = req.query.search || '';
        const page = parseInt(req.query.page) || 1;
        const limit = parseInt(req.query.limit) || 50;
        const sortBy = req.query.sortBy || '';
        const sortOrder = req.query.sortOrder || 'asc';
        const filters = req.query.filters ? JSON.parse(req.query.filters) : {};

        let filteredData = [...cleanedData];

        if (search) {
            const searchLower = search.toLowerCase();
            filteredData = filteredData.filter(row => {
                return Object.values(row).some(value => {
                    if (value === null || value === undefined) return false;
                    return String(value).toLowerCase().includes(searchLower);
                });
            });
        }

        Object.keys(filters).forEach(key => {
            const filter = filters[key];
            filteredData = filteredData.filter(row => {
                const value = row[key];
                if (filter.type === 'equals') return String(value) === String(filter.value);
                if (filter.type === 'contains') return String(value).toLowerCase().includes(String(filter.value).toLowerCase());
                if (filter.type === 'greater') return parseFloat(value) > parseFloat(filter.value);
                if (filter.type === 'less') return parseFloat(value) < parseFloat(filter.value);
                if (filter.type === 'between') {
                    const val = parseFloat(value);
                    return val >= parseFloat(filter.min) && val <= parseFloat(filter.max);
                }
                return true;
            });
        });

        if (sortBy && filteredData.length > 0) {
            filteredData.sort((a, b) => {
                const aVal = a[sortBy] || '';
                const bVal = b[sortBy] || '';
                if (sortOrder === 'asc') {
                    return String(aVal).localeCompare(String(bVal));
                } else {
                    return String(bVal).localeCompare(String(aVal));
                }
            });
        }

        const total = filteredData.length;
        const startIndex = (page - 1) * limit;
        const endIndex = Math.min(startIndex + limit, total);
        const paginatedData = filteredData.slice(startIndex, endIndex);

        const stats = calculateStats(filteredData);

        res.json({
            data: paginatedData,
            headers: currentHeaders,
            stats: stats,
            columnTypes: columnTypes,
            summary: dataSummary,
            total: total,
            page: page,
            limit: limit,
            totalPages: Math.ceil(total / limit)
        });
    } catch (error) {
        res.status(500).json({ error: 'Error fetching data: ' + error.message });
    }
});

// Clean Data
app.post('/api/clean', (req, res) => {
    try {
        const cleanOptions = req.body.options || {};
        cleanedData = cleanDataset(uploadedData, cleanOptions);
        
        const stats = calculateStats(cleanedData);
        dataHistory.push({
            id: uuidv4(),
            timestamp: new Date().toISOString(),
            action: 'clean',
            data: cleanedData,
            headers: currentHeaders
        });

        res.json({
            success: true,
            message: 'Data cleaned successfully',
            data: cleanedData,
            stats: stats,
            total: cleanedData.length
        });
    } catch (error) {
        res.status(500).json({ error: 'Error cleaning data: ' + error.message });
    }
});

// EDA (Exploratory Data Analysis)
app.get('/api/eda', (req, res) => {
    try {
        const edaResults = performEDA(cleanedData, columnTypes);
        res.json({
            success: true,
            eda: edaResults
        });
    } catch (error) {
        res.status(500).json({ error: 'Error performing EDA: ' + error.message });
    }
});

// Export to Excel
app.get('/api/export/excel', (req, res) => {
    try {
        const data = cleanedData.map(row => {
            const newRow = {};
            currentHeaders.forEach(h => {
                newRow[h] = row[h] || '';
            });
            return newRow;
        });

        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(data);
        XLSX.utils.book_append_sheet(wb, ws, 'Data');
        const filename = `export_${Date.now()}.xlsx`;
        const filepath = path.join(__dirname, filename);
        XLSX.writeFile(wb, filepath);

        res.download(filepath, filename, (err) => {
            fs.unlinkSync(filepath);
            if (err) console.error('Download error:', err);
        });
    } catch (error) {
        res.status(500).json({ error: 'Error exporting: ' + error.message });
    }
});

// Export to PDF
app.get('/api/export/pdf', (req, res) => {
    try {
        const doc = new PDFDocument();
        const filename = `report_${Date.now()}.pdf`;
        const filepath = path.join(__dirname, filename);
        const writeStream = fs.createWriteStream(filepath);

        doc.pipe(writeStream);
        
        doc.fontSize(20).text('Data Report - SAJID WORLD', { align: 'center' });
        doc.moveDown();
        doc.fontSize(12).text(`Generated: ${new Date().toLocaleString()}`);
        doc.text(`Total Records: ${cleanedData.length}`);
        doc.text(`Columns: ${currentHeaders.join(', ')}`);
        doc.moveDown();

        const stats = calculateStats(cleanedData);
        doc.fontSize(14).text('Data Statistics:');
        doc.fontSize(10);
        Object.keys(stats.columnStats).forEach(col => {
            const info = stats.columnStats[col];
            doc.text(`${col}: ${info.type} | Unique: ${info.unique} | Count: ${info.count}`);
            if (info.type === 'numeric') {
                doc.text(`  Min: ${info.min} | Max: ${info.max} | Avg: ${info.avg}`);
            }
        });

        doc.end();

        writeStream.on('finish', () => {
            res.download(filepath, filename, (err) => {
                fs.unlinkSync(filepath);
                if (err) console.error('Download error:', err);
            });
        });
    } catch (error) {
        res.status(500).json({ error: 'Error generating PDF: ' + error.message });
    }
});

// Delete rows
app.delete('/api/data', (req, res) => {
    try {
        const { ids } = req.body;
        if (!ids || !Array.isArray(ids)) {
            return res.status(400).json({ error: 'Invalid IDs provided' });
        }

        const idSet = new Set(ids);
        cleanedData = cleanedData.filter(row => !idSet.has(row.id));
        uploadedData = [...cleanedData];

        dataHistory.push({
            id: uuidv4(),
            timestamp: new Date().toISOString(),
            action: 'delete',
            data: cleanedData,
            headers: currentHeaders
        });

        const stats = calculateStats(cleanedData);
        res.json({
            success: true,
            message: `Deleted ${ids.length} row(s)`,
            data: cleanedData,
            stats: stats,
            total: cleanedData.length
        });
    } catch (error) {
        res.status(500).json({ error: 'Error deleting data: ' + error.message });
    }
});

// AI Insights
app.get('/api/ai/insights', (req, res) => {
    try {
        const insights = generateAIInsights(cleanedData, columnTypes);
        res.json({
            success: true,
            insights: insights
        });
    } catch (error) {
        res.status(500).json({ error: 'Error generating AI insights: ' + error.message });
    }
});

// Save Search
app.post('/api/save-search', (req, res) => {
    try {
        const { name, query } = req.body;
        savedSearches.push({
            id: uuidv4(),
            name: name,
            query: query,
            timestamp: new Date().toISOString()
        });
        res.json({
            success: true,
            message: 'Search saved successfully',
            savedSearches: savedSearches
        });
    } catch (error) {
        res.status(500).json({ error: 'Error saving search: ' + error.message });
    }
});

// Get Saved Searches
app.get('/api/saved-searches', (req, res) => {
    res.json({
        success: true,
        savedSearches: savedSearches
    });
});

// Reset
app.delete('/api/reset', (req, res) => {
    try {
        uploadedData = [];
        cleanedData = [];
        currentHeaders = [];
        dataId = null;
        columnTypes = {};
        dataSummary = {};
        res.json({
            success: true,
            message: 'Data reset successfully'
        });
    } catch (error) {
        res.status(500).json({ error: 'Error resetting data: ' + error.message });
    }
});

// ============ HELPER FUNCTIONS ============

function detectColumnTypes(data) {
    if (!data || data.length === 0) return {};
    const types = {};
    const sampleSize = Math.min(data.length, 100);
    const sample = data.slice(0, sampleSize);
    
    for (let key of Object.keys(data[0])) {
        if (key === 'id' || key === '_index') continue;
        
        let numericCount = 0, dateCount = 0, stringCount = 0;
        let uniqueValues = new Set();
        let totalCount = 0;
        
        for (let row of sample) {
            const val = row[key];
            if (val === null || val === undefined || val === '') continue;
            totalCount++;
            uniqueValues.add(String(val).trim());
            
            const numVal = parseFloat(val);
            if (!isNaN(numVal) && isFinite(val)) {
                numericCount++;
            } else if (!isNaN(Date.parse(val))) {
                dateCount++;
            } else {
                stringCount++;
            }
        }
        
        if (numericCount / totalCount > 0.6) {
            types[key] = { type: 'numeric', uniqueValues: uniqueValues.size, sampleCount: totalCount };
        } else if (dateCount / totalCount > 0.6) {
            types[key] = { type: 'date', uniqueValues: uniqueValues.size, sampleCount: totalCount };
        } else {
            types[key] = { type: 'categorical', uniqueValues: uniqueValues.size, sampleCount: totalCount };
        }
    }
    return types;
}

function generateDataSummary(data, types) {
    const summary = {};
    for (let key of Object.keys(data[0])) {
        if (key === 'id' || key === '_index') continue;
        const values = data.map(row => row[key]).filter(v => v !== null && v !== undefined && v !== '');
        const type = types[key]?.type || 'unknown';
        summary[key] = {
            type: type,
            count: values.length,
            unique: new Set(values.map(v => String(v).trim())).size,
            sample: values.slice(0, 5)
        };
        if (type === 'numeric') {
            const numValues = values.map(v => parseFloat(v)).filter(v => !isNaN(v));
            if (numValues.length > 0) {
                summary[key].min = Math.min(...numValues);
                summary[key].max = Math.max(...numValues);
                summary[key].avg = (numValues.reduce((a, b) => a + b, 0) / numValues.length).toFixed(2);
                summary[key].sum = numValues.reduce((a, b) => a + b, 0);
                summary[key].median = stats.median(numValues);
                summary[key].std = stats.sampleStandardDeviation(numValues);
            }
        }
    }
    return summary;
}

function calculateStats(data) {
    if (!data || data.length === 0) {
        return { rows: 0, cols: 0, totalRecords: 0, columnStats: {} };
    }

    const sample = data[0];
    const headers = Object.keys(sample).filter(h => h !== 'id' && h !== '_index');
    const columnStats = {};

    for (let header of headers) {
        const values = data.map(row => row[header]).filter(v => v !== null && v !== undefined && v !== '');
        const type = columnTypes[header]?.type || 'categorical';
        columnStats[header] = {
            type: type,
            count: values.length,
            unique: new Set(values.map(v => String(v).trim())).size
        };
        if (type === 'numeric') {
            const numValues = values.map(v => parseFloat(v)).filter(v => !isNaN(v));
            if (numValues.length > 0) {
                columnStats[header].min = Math.min(...numValues);
                columnStats[header].max = Math.max(...numValues);
                columnStats[header].avg = (numValues.reduce((a, b) => a + b, 0) / numValues.length).toFixed(2);
                columnStats[header].sum = numValues.reduce((a, b) => a + b, 0);
                columnStats[header].median = stats.median(numValues);
                columnStats[header].std = stats.sampleStandardDeviation(numValues);
            }
        }
    }

    return {
        rows: data.length,
        cols: headers.length,
        totalRecords: data.length,
        columnStats: columnStats
    };
}

function cleanDataset(data, options) {
    let cleaned = [...data];
    
    if (options.removeDuplicates) {
        const seen = new Set();
        cleaned = cleaned.filter(row => {
            const key = currentHeaders.map(h => row[h]).join('|');
            if (seen.has(key)) return false;
            seen.add(key);
            return true;
        });
    }

    if (options.removeEmpty) {
        cleaned = cleaned.filter(row => {
            return Object.values(row).some(v => v !== null && v !== undefined && v !== '');
        });
    }

    if (options.fillMissing) {
        cleaned = cleaned.map(row => {
            currentHeaders.forEach(h => {
                if (row[h] === null || row[h] === undefined || row[h] === '') {
                    if (columnTypes[h]?.type === 'numeric') {
                        const values = cleaned.map(r => parseFloat(r[h])).filter(v => !isNaN(v));
                        row[h] = values.length > 0 ? stats.mean(values) : 0;
                    } else {
                        row[h] = 'N/A';
                    }
                }
            });
            return row;
        });
    }

    if (options.removeOutliers) {
        const numericCols = Object.keys(columnTypes).filter(k => columnTypes[k]?.type === 'numeric');
        numericCols.forEach(col => {
            const values = cleaned.map(r => parseFloat(r[col])).filter(v => !isNaN(v));
            if (values.length > 0) {
                const q1 = stats.quantile(values, 0.25);
                const q3 = stats.quantile(values, 0.75);
                const iqr = q3 - q1;
                const lowerBound = q1 - 1.5 * iqr;
                const upperBound = q3 + 1.5 * iqr;
                cleaned = cleaned.filter(row => {
                    const val = parseFloat(row[col]);
                    return isNaN(val) || (val >= lowerBound && val <= upperBound);
                });
            }
        });
    }

    return cleaned;
}

function performEDA(data, types) {
    const results = {
        summary: {},
        correlations: {},
        distributions: {},
        anomalies: [],
        insights: []
    };

    const numericCols = Object.keys(types).filter(k => types[k]?.type === 'numeric');
    const categoricalCols = Object.keys(types).filter(k => types[k]?.type === 'categorical');

    numericCols.forEach(col => {
        const values = data.map(r => parseFloat(r[col])).filter(v => !isNaN(v));
        if (values.length > 0) {
            results.summary[col] = {
                count: values.length,
                mean: stats.mean(values),
                median: stats.median(values),
                mode: stats.mode(values),
                min: Math.min(...values),
                max: Math.max(...values),
                variance: stats.variance(values),
                std: stats.sampleStandardDeviation(values),
                skewness: stats.skewness(values),
                kurtosis: stats.kurtosis(values)
            };
        }
    });

    for (let i = 0; i < numericCols.length; i++) {
        for (let j = i + 1; j < numericCols.length; j++) {
            const col1 = numericCols[i];
            const col2 = numericCols[j];
            const pairs = data
                .map(row => [parseFloat(row[col1]), parseFloat(row[col2])])
                .filter(([a, b]) => !isNaN(a) && !isNaN(b));
            if (pairs.length > 1) {
                const corr = calculateCorrelation(pairs);
                results.correlations[`${col1} vs ${col2}`] = corr;
                if (Math.abs(corr) > 0.7) {
                    results.insights.push(`Strong ${corr > 0 ? 'positive' : 'negative'} correlation between ${col1} and ${col2} (${corr.toFixed(3)})`);
                }
            }
        }
    }

    categoricalCols.forEach(col => {
        const dist = {};
        data.forEach(row => {
            const val = String(row[col] || 'unknown').trim();
            dist[val] = (dist[val] || 0) + 1;
        });
        results.distributions[col] = Object.entries(dist).sort((a, b) => b[1] - a[1]);
    });

    numericCols.forEach(col => {
        const values = data.map(r => parseFloat(r[col])).filter(v => !isNaN(v));
        if (values.length > 0) {
            const mean = stats.mean(values);
            const std = stats.sampleStandardDeviation(values);
            const threshold = 3 * std;
            data.forEach((row, idx) => {
                const val = parseFloat(row[col]);
                if (!isNaN(val) && Math.abs(val - mean) > threshold) {
                    results.anomalies.push({
                        row: idx,
                        column: col,
                        value: val,
                        deviation: val - mean
                    });
                }
            });
        }
    });

    return results;
}

function generateAIInsights(data, types) {
    const insights = {
        summary: [],
        patterns: [],
        recommendations: [],
        anomalies: []
    };

    insights.summary.push(`Dataset contains ${data.length} records with ${Object.keys(types).length} columns`);
    
    const numericCols = Object.keys(types).filter(k => types[k]?.type === 'numeric');
    const categoricalCols = Object.keys(types).filter(k => types[k]?.type === 'categorical');
    
    insights.summary.push(`Found ${numericCols.length} numeric columns and ${categoricalCols.length} categorical columns`);

    categoricalCols.forEach(col => {
        const dist = {};
        data.forEach(row => {
            const val = String(row[col] || 'unknown').trim();
            dist[val] = (dist[val] || 0) + 1;
        });
        const sorted = Object.entries(dist).sort((a, b) => b[1] - a[1]);
        if (sorted.length > 0) {
            const top = sorted[0];
            insights.patterns.push(`In "${col}", "${top[0]}" appears most frequently (${top[1]} times, ${((top[1]/data.length)*100).toFixed(1)}%)`);
        }
    });

    numericCols.forEach(col => {
        const values = data.map(r => parseFloat(r[col])).filter(v => !isNaN(v));
        if (values.length > 0) {
            const mean = stats.mean(values);
            insights.recommendations.push(`Consider analyzing "${col}" trends - average is ${mean.toFixed(2)}`);
        }
    });

    numericCols.forEach(col => {
        const values = data.map(r => parseFloat(r[col])).filter(v => !isNaN(v));
        if (values.length > 0) {
            const mean = stats.mean(values);
            const std = stats.sampleStandardDeviation(values);
            const threshold = 3 * std;
            let anomalies = 0;
            data.forEach(row => {
                const val = parseFloat(row[col]);
                if (!isNaN(val) && Math.abs(val - mean) > threshold) {
                    anomalies++;
                }
            });
            if (anomalies > 0) {
                insights.anomalies.push(`Found ${anomalies} potential outliers in "${col}"`);
            }
        }
    });

    return insights;
}

function calculateCorrelation(pairs) {
    const n = pairs.length;
    if (n < 2) return 0;
    const sumX = pairs.reduce((s, [x]) => s + x, 0);
    const sumY = pairs.reduce((s, [, y]) => s + y, 0);
    const sumXY = pairs.reduce((s, [x, y]) => s + x * y, 0);
    const sumX2 = pairs.reduce((s, [x]) => s + x * x, 0);
    const sumY2 = pairs.reduce((s, [, y]) => s + y * y, 0);
    const numerator = n * sumXY - sumX * sumY;
    const denominator = Math.sqrt((n * sumX2 - sumX * sumX) * (n * sumY2 - sumY * sumY));
    if (denominator === 0) return 0;
    return numerator / denominator;
}

app.use((err, req, res, next) => {
    console.error(err.stack);
    res.status(500).json({ error: err.message || 'Something went wrong!' });
});

function findAvailablePort(startPort, callback) {
    const server = app.listen(startPort, () => {
        server.close(() => callback(startPort));
    });
    server.on('error', (err) => {
        if (err.code === 'EADDRINUSE') {
            findAvailablePort(startPort + 1, callback);
        } else {
            console.error('Error:', err);
            process.exit(1);
        }
    });
}

findAvailablePort(DEFAULT_PORT, (port) => {
    app.listen(port, () => {
        console.log(`🚀 AUTHENTIC DATA DASHBOARD by SAJID WORLD`);
        console.log(`📍 Running on http://localhost:${port}`);
        console.log(`📊 Universal Data Dashboard - Works with ANY dataset!`);
        console.log(`💡 Features: Upload, Clean, EDA, AI Insights, Export, Zoom, and more!`);
    });
});
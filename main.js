const {app, BrowserWindow, Tray, Menu } = require('electron');
const path = require('path');
const express = require('express');
const { createProxyMiddleware } = require('http-proxy-middleware');
const ExporterService = require('./ExpoterService');

// Electron: Prevent window from showing
let tray = null;
app.on('ready', () => {
    const win = new BrowserWindow({
        show: false, // Prevent window from being shown
        webPreferences: {
            nodeIntegration: true,
        }
    });
    // Set the tray icon
    tray = new Tray(path.join(__dirname, 'icon.ico'));
    const contextMenu = Menu.buildFromTemplate([
        { label: 'Quit', click: () => app.quit() }
    ]);

    tray.setToolTip('ExportExcelApp');
    tray.setContextMenu(contextMenu);

    // Here, run your Node.js server
    startExpressApp();
});

function startExpressApp() {
    const expressApp = express(); // Renamed to avoid conflict with Electron's app
    expressApp.use(express.json());
    expressApp.use(express.urlencoded({ extended: true }));

    // Route that handles /xuatxlsx
    expressApp.get('/xuatxlsx', async (req, res) => {
        var filename = req.query.tenfile;
        var mau = req.query.mau;

        try {
            var jtt = JSON.parse(customDecode(req.query.jtt));
            var jct = JSON.parse(customDecode(req.query.jct).replace(/(\r\n|\n|\r|\t)/gm, " "));
            var jft = JSON.parse(customDecode(req.query.jft));

            console.log(jtt);
            console.log(jft);
            console.log(jct);


            jtt = JSON.parse(JSON.stringify(jtt[0]));
            var jttMap = new Map(Object.entries(jtt));

            var listJctMap = [];
            for (let i = 0; i < jct.length; i++) {
                listJctMap.push(new Map(Object.entries(JSON.parse(JSON.stringify(jct[i])))));
            }

            const arrayFromMaps = listJctMap.map(map => Object.fromEntries(map));
            const sortedArray = arrayFromMaps.sort((a, b) => {
                const aValue = a['A'] || '';
                const bValue = b['A'] || '';
                return aValue.localeCompare(bValue, undefined, { numeric: true });
            });
            const sortedListJctMap = sortedArray.map(obj => new Map(Object.entries(obj)));

            var jftMap = new Map(Object.entries(JSON.parse(JSON.stringify(jft[0]))));

            const exporterService = new ExporterService(filename, mau, jttMap, sortedListJctMap, jftMap);
            await exporterService.exportToExcel(res);
        } catch (error) {
            console.error(error);
            res.status(500).send('Error exporting Excel file');
        }
    });

    const options = {
        target: 'http://localhost:3000',
        changeOrigin: true,
        pathRewrite: {
            '^/proxy/xuatxlsx': '/xuatxlsx',
        },
    };

    const exampleProxy = createProxyMiddleware(options);
    expressApp.use('/proxy/xuatxlsx', exampleProxy);

    const PORT = 80;
    expressApp.listen(PORT, () => {
        console.log(`Server is running on port ${PORT}`);
    });
}

// Regular expression to match %X where X is not in the range 2-7
//const Pattern = /%[%)}\[{*&^$#@]/g;
const Pattern = /%[%()}\[{*&^$#@a-zA-Z]/g;
function sanitizeUrl(url) {
    return url.replace(Pattern, (match) => {
        const invalidEncoding = match.slice(1);
        return `%25${invalidEncoding}`;
    });
}

function customDecode(str) {
    if (Pattern.test(str)) {
        var value = sanitizeUrl(str);
        var decoded = decodeURIComponent(value);
        return decoded;
    } else {
        return str;
    }
}

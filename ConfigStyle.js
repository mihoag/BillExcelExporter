const fs = require('fs');
const path = require('path');
const ConfigEntry = require('./ConfigEntry'); // Assuming ConfigEntry is in the same directory

const filepath = path.join(__dirname, 'config.txt');

function getConfig() {
    const configEntries = {};
    let currentEntry = null;
    let currentSection = null;

    const lines = fs.readFileSync(filepath, 'utf-8').split('\n');
    
    for (let line of lines) {
        line = line.trim();
        if (line === '' || line.startsWith('#')) continue;

        if (line.startsWith('[')) {
            if (currentSection && currentEntry) {
                configEntries[currentSection] = currentEntry;
            }
            currentSection = line.substring(1, line.length - 1).trim();
            currentEntry = new ConfigEntry();
        } else if (line.includes('=')) {
            const [key, value] = line.split('=', 2).map(str => str.trim());

            switch (key) {
                case 'header_row':
                    currentEntry.setHeaderRow(parseInt(value));
                    break;
                case 'alignment':
                    currentEntry.setAlignment(parseMap(value));
                    break;
                case 'format_number':
                    currentEntry.setFormatNumber(parseList(value));
                    break;
                case 'font_color':
                    currentEntry.setFontColor(parseMap(value));
                    break;
                case 'formula_multiply':
                    currentEntry.setFormulaMultiply(value.split(',').map(str => str.trim()));
                    break;
            }
        }
    }
    
    if (currentSection && currentEntry) {
        configEntries[currentSection] = currentEntry;
    }

    return configEntries;
}

function parseMap(value) {
    return value.substring(1, value.length - 1).split(',').reduce((map, entry) => {
        const [key, val] = entry.split(':').map(str => str.trim().replace(/"/g, ''));
        map[key] = val;
        return map;
    }, {});
}

function parseList(value) {
    return value.substring(1, value.length - 1).split(',').map(str => str.trim());
}
module.exports = {
    getConfig
};


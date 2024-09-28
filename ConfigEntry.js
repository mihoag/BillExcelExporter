class ConfigEntry {
    constructor() {
        this.headerRow = 0;
        this.alignment = {};
        this.format_number = [];
        this.fontColor = {};
        this.formulaMultiply = [];
    }

    setHeaderRow(value) {
        this.headerRow = value;
    }

    setAlignment(value) {
        this.alignment = value;
    }

    setFormatNumber(value) {
        this.format_number = value;
    }

    setFontColor(value) {
        this.fontColor = value;
    }

    setFormulaMultiply(value) {
        this.formulaMultiply = value;
    }
}

module.exports = ConfigEntry;

class WorkSheets {
    sheets: Sheet[];
    constructor() {
        this.sheets = [];
    }
}

class Sheet {
    name: string;
    columns: string[];
    constructor(name: string, columns: string[]) {
        this.name = name;
        this.columns = columns;
    }
}

module.exports = {
    WorkSheets: WorkSheets,
    Sheet: Sheet
};
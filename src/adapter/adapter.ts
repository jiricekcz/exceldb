export interface CellTypeDefault {
    storageType: string | number | boolean | Date;
}
export interface CellTypeVirtual {
    storageType: string | number | boolean | Date;
    virtualType: any;
}
export type CellType = CellTypeDefault | CellTypeVirtual;
export type SheetStructure = Record<string, CellType>;
export type WorkbookStructure = Record<string, SheetStructure>;

export type StorageTypesTypeMap = {
    "string": string;
    "number": number;
    "boolean": boolean;
    "Date": Date;
}
export type StringFromType<T extends StorageTypesTypeMap[keyof StorageTypesTypeMap]> = T extends StorageTypesTypeMap[infer R extends keyof StorageTypesTypeMap] ? R : never;

export interface CellConstructorOptionsDefault<T extends CellTypeDefault> {
    storageType: StringFromType<T["storageType"]>;
    validators?: {
        [name: string]: (value: T["storageType"]) => boolean
    };
} 

export interface CellConstructorOptionsVirtual<T extends CellTypeVirtual> extends CellConstructorOptionsDefault<T> {
    getter: (raw: T["storageType"]) => T["virtualType"];
    setter: (virtual: T["virtualType"]) => T["storageType"];
    validators?: {
        [name: string]: (value: T["virtualType"]) => boolean
    }
}
export type CellInterfaceType<C extends CellType> = C extends CellTypeVirtual ? C["virtualType"] : C["storageType"]; 
export type CellConstructorOptions<T extends CellType> = T extends CellTypeVirtual ? CellConstructorOptionsVirtual<T> : CellConstructorOptionsDefault<T>; 
export type WorkbookConstructorOptions<T extends WorkbookStructure> = {
    sheets: {
        [SheetName in keyof T]: {
            columns: {
                [CollName in keyof T[SheetName]]: CellConstructorOptions<T[SheetName][CollName]>;
            }
        }
    }
}


export type AssertExtends<M, D extends M> = D;

export abstract class WorkbookAdapter<W extends WorkbookStructure> {
    abstract getSheetNames(): (keyof W)[];
    abstract getSheet<S extends keyof W>(sheetName: S): Sheet<W, S>;
    protected* getSheets() {
        for (const sheetName of this.getSheetNames()) {
            yield this.getSheet(sheetName);
        }
    }
    get sheets() {
        return this.getSheets();
    }
    [Symbol.iterator]() {
        return this.getSheets();
    }
}

export abstract class Sheet<W extends WorkbookStructure, N extends keyof W> {
    public readonly wokrbook: WorkbookAdapter<W>;
    public readonly name: N;
    constructor(workbook: WorkbookAdapter<W>, sheetName: N) {
        this.wokrbook = workbook;
        this.name = sheetName;
    }

    public abstract getRowCount(): number;
    public abstract getRow(rowNumber: number): Row<W, N>;

    public abstract getColumnNames(): (keyof W[N])[];
    public abstract getColumn<C extends keyof W[N]>(columnName: C): Column<W, N, C>;

    protected *getRows() {
        const rowCount = this.getRowCount();
        for (let i = 0; i < rowCount; i++) {
            yield this.getRow(i);
        }
    }

    protected *getColumns() {
        for (const columnName of this.getColumnNames()) {
            yield this.getColumn(columnName);
        }
    }

    get rows() {
        return this.getRows();
    }

    get columns() {
        return this.getColumns();
    }

    [Symbol.iterator]() {
        return this.getRows();
    }
}

export abstract class Column<W extends WorkbookStructure, S extends keyof W, N extends keyof W[S]> {
    public readonly workbook: WorkbookAdapter<W>;
    public readonly sheet: Sheet<W, S>;
    public readonly name: N;
    constructor(sheet: Sheet<W, S>, columnName: N) {
        this.sheet = sheet;
        this.name = columnName;
        this.workbook = sheet.wokrbook;
    }
}



export abstract class Row<W extends WorkbookStructure, S extends keyof W> {
    public readonly workbook: WorkbookAdapter<W>;
    public readonly sheet: Sheet<W, S>;
    public rowNumber: number;

    public readonly cells: {
        [C in keyof W[S]]: CellInterfaceType<W[S][C]>
    };
    constructor(sheet: Sheet<W, S>, rowNumber: number) {
        this.sheet = sheet;
        this.rowNumber = rowNumber;
        this.workbook = sheet.wokrbook;

        const cells: {[C in keyof W[S]]?: CellInterfaceType<W[S][C]>} = {};

        for (const columnName of sheet.getColumnNames()) {
            Object.defineProperty(cells, columnName, {
                get: () => this.getCellValue(columnName),
                set: (value) => this.setCellValue(columnName, value),
                enumerable: true
            });
        }            
        this.cells = cells as {[C in keyof W[S]]: CellInterfaceType<W[S][C]>};
    }
    public abstract getCellValue<C extends keyof W[S]>(columnName: C): CellInterfaceType<W[S][C]>;
    public abstract setCellValue<C extends keyof W[S]>(columnName: C, value: CellInterfaceType<W[S][C]>): void;


}

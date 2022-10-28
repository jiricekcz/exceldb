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

export type StorageTypes = [
    ["string", string],
    ["number", number],
    ["boolean", boolean],
    ["Date", Date],
];

export type StorageTypeString = StorageTypes[number][0];
export type StorageTypeType = StorageTypes[number][1];


export type StorageTypeToType<S extends StorageTypeString> = S extends StorageTypes[0][0] ? StorageTypes[0][1] : S extends StorageTypes[1][0] ? StorageTypes[1][1] : S extends StorageTypes[2][0] ? StorageTypes[2][1] : S extends StorageTypes[3][0] ? StorageTypes[3][1] : never;
export type TypeToStorageType<T extends StorageTypeType> = T extends StorageTypes[0][1] ? StorageTypes[0][0] : T extends StorageTypes[1][1] ? StorageTypes[1][0] : T extends StorageTypes[2][1] ? StorageTypes[2][0] : T extends StorageTypes[3][1] ? StorageTypes[3][0] : never;
// export type StringFromType<T extends StorageTypesTypeMap[keyof StorageTypesTypeMap]> = T extends StorageTypesTypeMap[infer R] ? R : never;

export interface CellConstructorOptionsDefault<T extends CellTypeDefault> {
    storageType: TypeToStorageType<T["storageType"]>;
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
    protected abstract generateFile(): Promise<string | Uint8Array | Buffer>;
    protected abstract saveFile(): Promise<void>;
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

    protected virtualCellCache: {
        [sheetname in keyof W]?: {
            [cellname in keyof W[sheetname]]?: {
                [rowNumber: number]: CellInterfaceType<W[sheetname][cellname]> | undefined;
            }
        }
    } = {};
    public getVirtualCellCache<S extends keyof W, C extends keyof W[S]>(sheetname: S, columnname: C, rowNumber: number): CellInterfaceType<W[S][C]> | undefined {
        if (!(sheetname in this.virtualCellCache)) this.virtualCellCache[sheetname] = {}; // If the sheat doesn't exist in cache, create it
        if (!(columnname in (this.virtualCellCache[sheetname] as any))) (this.virtualCellCache[sheetname] as any)[columnname] = {}; // If the column doesn't exist in cache, create it
        return (this.virtualCellCache[sheetname] as any)[columnname][rowNumber]; // Return the value

    }

    public setVirtualCellCache<S extends keyof W, C extends keyof W[S]>(sheetname: S, cellname: C, rowNumber: number, value: CellInterfaceType<W[S][C]>): void {
        if (!(sheetname in this.virtualCellCache)) this.virtualCellCache[sheetname] = {}; // If the sheat doesn't exist in cache, create it
        if (!(cellname in (this.virtualCellCache[sheetname] as any))) (this.virtualCellCache[sheetname] as any)[cellname] = {}; // If the column doesn't exist in cache, create it
        const v = (this.virtualCellCache[sheetname] as any)[cellname][rowNumber]; // Get the previous value
        if (v !== undefined && typeof v === "object" && v !== null && v !== value) { // If the previous value is an object and it's not the same as the new value
            for (const key in v) {
                Object.defineProperty(v, key, { get: () => {
                    throw new Error("Cannot modify or read virtual cell object after it has been overwritten. This is to prevent writing to old virtual cells, that are no longer bound to the databse saves.");
                },
                set: () => {
                    throw new Error("Cannot modify or read virtual cell object after it has been overwritten. This is to prevent writing to old virtual cells, that are no longer bound to the databse saves.");
                }
             }); // Propetry access will throw an error to prevent accidental modification to no longer binded object
            }
        }
        
        (this.virtualCellCache[sheetname] as any)[cellname][rowNumber] = value;
    }
    public save(): Promise<void> {
        this.preSave();
        return this.saveFile();
    }
    public export(): Promise<string | Uint8Array | Buffer>{
        this.preSave();
        return this.generateFile();
    }
    public preSave(): void {
        for (const sheet of this) {
            sheet.preSave();
        }
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


    public getVirtualCellCache(rowNumber: number, columnName: keyof W[N]): undefined | CellInterfaceType<W[N][keyof W[N]]> {
        return this.wokrbook.getVirtualCellCache(this.name, columnName, rowNumber);
    }

    public setVirtualCellCache(rowNumber: number, columnName: keyof W[N], value: CellInterfaceType<W[N][keyof W[N]]>): void {
        this.wokrbook.setVirtualCellCache(this.name, columnName, rowNumber, value);
    }

    public preSave(): void {
        for (const row of this.rows) {
            row.preSave();
        }
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
            const get = this.isCellVirtual(columnName) ? () => {
                let v = this.sheet.getVirtualCellCache(this.rowNumber, columnName);
                if (v === undefined) {
                    v = this.getCellValue(columnName);
                    this.sheet.setVirtualCellCache(this.rowNumber, columnName, v as CellInterfaceType<W[S][keyof W[S]]>);
                }
                return v;
            } : () => {
                return this.getCellValue(columnName);
            };
            const set = this.isCellVirtual(columnName) ? (value: CellInterfaceType<W[S][keyof W[S]]>) => {
                this.sheet.setVirtualCellCache(this.rowNumber, columnName, value);
            } : (value: CellInterfaceType<W[S][keyof W[S]]>) => this.setCellValue(columnName, value);
            Object.defineProperty(cells, columnName, {
                get,
                set,
                enumerable: true
            });
        }            
        this.cells = cells as {[C in keyof W[S]]: CellInterfaceType<W[S][C]>};
    }
    public abstract getCellValue<C extends keyof W[S]>(columnName: C): CellInterfaceType<W[S][C]>;
    public abstract setCellValue<C extends keyof W[S]>(columnName: C, value: CellInterfaceType<W[S][C]>): void;
    public abstract isCellVirtual<C extends keyof W[S]>(columnName: C): boolean;

    public preSave(): void {
        for (const columnName of this.sheet.getColumnNames()) {
            this.setCellValue(columnName, this.sheet.getVirtualCellCache(this.rowNumber, columnName) as CellInterfaceType<W[S][keyof W[S]]>);
        }
    }

}

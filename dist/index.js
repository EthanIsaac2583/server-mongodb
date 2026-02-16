/******/ (() => { // webpackBootstrap
/******/ 	var __webpack_modules__ = ({

/***/ 44
(module, __unused_webpack_exports, __webpack_require__) {

const uuid = __webpack_require__(333);
const jwt = __webpack_require__(127);
const { MongoClient } = __webpack_require__(884);

const { setWidth, hideColumn, showColumn, insertColumn, moveColumn, deleteColumn } = __webpack_require__(802);
const { setHeight, hideRow, showRow, moveRow, insertRow, deleteRow, setRowId } = __webpack_require__(954);
const { setMeta, resetMeta } = __webpack_require__(657);
const { setComments } = __webpack_require__(998);
const { setFreezeColumns } = __webpack_require__(348);
const { setFreezeRows } = __webpack_require__(530);
const { setFooter, setFooterValue, resetFooter } = __webpack_require__(295);
const { setHeader } = __webpack_require__(701);
const { setNestedHeaders, resetNestedHeaders, setNestedCell } = __webpack_require__(349);
const { setMerge, removeMerge } = __webpack_require__(562);
const { setCache } = __webpack_require__(404);
const { setStyle, resetStyle } = __webpack_require__(315);
const { setProperty, updateProperty } = __webpack_require__(381);
const { setValidations } = __webpack_require__(774);
const { setMedia } = __webpack_require__(502);
const { createWorksheet, deleteWorksheet, renameWorksheet, moveWorksheet,setWorksheetState, validateWorksheet } = __webpack_require__(666);
const { setValue, setFormula } = __webpack_require__(736);
const { setColumnGroup, setRowGroup } = __webpack_require__(485);
const { setDefinedNames } = __webpack_require__(227);
const { setConfig } = __webpack_require__(194);
const { orderBy } = __webpack_require__(574);

;(function(global, factory) {
     true ? module.exports = factory() :
    0;
}(this, (function() {

    const url = process.env.MONGODB_URI || process.env.MONGODB_URL || 'mongodb://mongodb';
    const dbName = process.env.MONGODB_DB_NAME || 'jspreadsheet';
    const collectionName = process.env.MONGODB_COLLECTION_NAME || 'documents';

    console.log('here -------------->', { url, dbName, collectionName })

    const client = new MongoClient(url);

    client.connect();

    const db = client.db(dbName);
    const collection = db.collection(collectionName);

    const methods = {
        setConfig,
        setWidth,
        hideColumn,
        showColumn,
        insertColumn,
        moveColumn,
        deleteColumn,
        setHeight,
        hideRow,
        showRow,
        moveRow,
        setMeta,
        resetMeta,
        setComments,
        setFreezeColumns,
        setFreezeRows,
        setFooter,
        setFooterValue,
        resetFooter,
        setHeader,
        setMerge,
        removeMerge,
        setColumnGroup,
        setRowGroup,
        setNestedHeaders,
        setNestedCell,
        resetNestedHeaders,
        setCache,
        setValidations,
        setStyle,
        resetStyle,
        setProperty,
        updateProperty,
        setMedia,
        createWorksheet,
        deleteWorksheet,
        renameWorksheet,
        moveWorksheet,
        setWorksheetState,
        setValue,
        setFormula,
        setDefinedNames,
        insertRow,
        deleteRow,
        setRowId,
        orderBy,
    }

    const getUserId = function(query) {
        // Decode token
        let user_id = null;
        if (query && query.token) {
            let info = jwt.decode(query.token);
            if (info) {
                user_id = info.sub;
            }
        }
        return user_id;
    }

    const get = async function(guid, query) {
        let config = await collection.findOne({ _id: guid });
        if (! config) {
            return false;
        }
        return config;
    }

    const load = async function(guid, query) {
        let config = await get(guid, query);
        if (! config) {
            return false;
        }
        // Make sure the owner is sent to the frontend
        config.spreadsheet.user_id = config.user_id
        return config.spreadsheet;
    }

    const create = async function(guid, config, query) {
        if (guid) {
            const result = await get(guid)
            if (result) {
                return { success: 1 };
            }
        } else {
            guid = uuid.v4();
        }

        // Decode token
        let user_id = getUserId(query);

        if (! config.style) {
            config.style = [];
        }
        if (! config.validations) {
            config.validations = [];
        }

        const worksheets = config.worksheets;
        const worksheetsLength = worksheets.length;
        for (let worksheetIndex = 0; worksheetIndex < worksheetsLength; worksheetIndex++) {
            const worksheet = worksheets[worksheetIndex];

            validateWorksheet(worksheet);
        }

        // Create a new spreadsheet
        await collection.insertOne({
            _id: guid,
            status: 1,
            user_id: user_id,
            created: new Date().toISOString(),
            updated: new Date().toISOString(),
            spreadsheet: config,
        });

        return { success: 1, guid: guid };
    }

    const destroy = async function(guid) {
        const result = await collection.deleteOne({ _id: guid });
        if (result.deletedCount === 1) {
            return { success: 1 };
        } else {
            return { error: 1 };
        }
    }

    // Queue per guid to ensure sequential execution
    const queues = new Map();

    const processQueue = async function(guid) {
        const queue = queues.get(guid);
        if (!queue || queue.processing || queue.items.length === 0) {
            if (queue && !queue.processing && queue.items.length === 0) {
                queues.delete(guid);
            }
            return;
        }

        queue.processing = true;
        const item = queue.items.shift();

        const result = await executeChange(guid, item.obj, item.query, item.onerror);

        item.resolve();

        if (result.error) {
            // Clear remaining items from the queue
            const remainingItems = queue.items.splice(0);
            remainingItems.forEach(i => i.resolve());
            // Sync the instance state to the database
            const config = item.obj.instance.getConfig();
            await replace(guid, config);
        }

        queue.processing = false;

        if (queue.items.length === 0) {
            queues.delete(guid);
        } else {
            processQueue(guid);
        }
    }

    const executeChange = async function(guid, obj, query, onerror) {
        try {
            let method = obj.method;
            // Get updates
            if (methods[method]) {
                // Current worksheet index
                obj.worksheetIndex = obj.instance.worksheets.findIndex(worksheet => worksheet.options.worksheetId === obj.worksheet)
                // Current document
                obj.document = await collection.findOne({ _id: guid });
                // Execute controller
                let changes = methods[method](obj);
                // Any changes to be executed
                if (changes) {
                    if (Array.isArray(changes)) {
                        // Bulk updates
                        changes = changes.map(({$filter = {}, ...change}) => {
                            return {
                                updateOne: {
                                    filter: { _id: guid, ...$filter },
                                    update: change
                                }
                            };
                        });
                        await collection.bulkWrite(changes);
                    } else {
                        await collection.updateOne({ _id: guid }, changes);
                    }
                }
            }
            return { success: 1 };
        } catch (error) {
            console.error(error);

            if (typeof(onerror) === 'function') {
                onerror(error);
            }

            return { error: 1, message: error };
        }
    }

    const change = function(guid, obj, query, onerror) {
        return new Promise((resolve) => {
            if (!queues.has(guid)) {
                queues.set(guid, { processing: false, items: [] });
            }

            const queue = queues.get(guid);
            queue.items.push({ obj, query, onerror, resolve });
            processQueue(guid);
        });
    }

    const replace = async function(guid, config, query, onerror) {
        try {
            await collection.updateOne({ _id: guid }, { $set: { spreadsheet: config } });
            return { success: 1 };
        } catch (error) {
            console.error(error);

            if (typeof(onerror) === 'function') {
                onerror(error);
            }

            return { error: 1, message: error };
        }
    }

    const list = async function(query) {
        // Decode token to get user ID
        let user_id = getUserId(query);
        // Get all documents for the user with the specified fields
        const cursor = await collection.find({ user_id: user_id }, { projection: { guid: 1, 'spreadsheet.name': 1, updated: 1, 'spreadsheet.privacy': 1 } });
        // Set the result
        let sheets = await cursor.toArray();
        if (sheets.length) {
            sheets.forEach((v) => {
                v.guid = v._id;
                v.privacy = v.spreadsheet.privacy ? 1 : 0;
                v.name = v.spreadsheet.name;
                delete v.spreadsheet;
            });
        }
        return sheets;
    }

    const setUsers = async function(guid, data) {
        return await collection.updateOne({ _id: guid }, {$set: { users: data }});
    }

    const getUsers = async function(guid) {
        let result = await collection.findOne({ _id: guid }, { projection: { users: 1 } });
        return result.users || [];
    }

    const setPrompts = async function(guid, data) {
        return await collection.updateOne({ _id: guid }, {$set: { prompts: data }});
    }

    const getPrompts = async function(guid) {
        let result = await collection.findOne({ _id: guid }, { projection: { prompts: 1 } });
        return result.prompts || [];
    }


    /**
     * Create a plugin object
     */
    let Extension = function(options) {
    }

    Extension.get = get;
    Extension.list = list;
    Extension.load = load;
    Extension.create = create;
    Extension.destroy = destroy;
    Extension.change = change;
    Extension.replace = replace;
    Extension.getUsers = getUsers;
    Extension.setUsers = setUsers;
    Extension.getPrompts = getPrompts;
    Extension.setPrompts = setPrompts;

    return Extension;

})));


/***/ },

/***/ 404
(module) {

const setCache = function(obj) {
    let cacheObject = obj.args[0];
    let changes = { $set: {} };
    Object.keys(cacheObject).forEach((cell) => {
        changes.$set[`spreadsheet.worksheets.${obj.worksheetIndex}.cache.${cell}`] = cacheObject[cell];
    });
    return changes;
}

module.exports = {
    setCache,
}


/***/ },

/***/ 802
(module, __unused_webpack_exports, __webpack_require__) {

const { updateReferences, moveArrayItems, dataIsMatrix } = __webpack_require__(191);

const setWidth = function(obj) {
    let [columns, widths] = obj.args;

    if (! Array.isArray(columns)) {
        columns = [columns];
    }

    const columnsPath = `spreadsheet.worksheets.${obj.worksheetIndex}.columns`;

    const savedColumns = obj.document.spreadsheet.worksheets[obj.worksheetIndex].columns;

    const widthsIsAnArray = Array.isArray(widths);

    const set = {};

    const columnsLength = columns.length;
    for (let index = 0; index < columnsLength; index++) {
        const column = columns[index];

        const columnPath = `${columnsPath}.${column}`;

        const width = widthsIsAnArray ? widths[index] : widths;

        if (savedColumns[column]) {
            set[`${columnPath}.width`] = width;
        } else {
            set[columnPath] = {
                width
            };
        }
    }

    return {
        $set: set
    };
}

const visibility = function(obj, state) {
    let columns = obj.args[0];

    if (! Array.isArray(columns)) {
        columns = [columns];
    }

    const columnsPath = `spreadsheet.worksheets.${obj.worksheetIndex}.columns`;

    const savedColumns = obj.document.spreadsheet.worksheets[obj.worksheetIndex].columns;

    const set = {};

    const columnsLength = columns.length;
    for (let i = 0; i < columnsLength; i++) {
        const column = columns[i];

        const columnPath = `${columnsPath}.${column}`;

        if (savedColumns[column]) {
            set[`${columnPath}.visible`] = state;
        } else {
            set[columnPath] = {
                visible: state
            };
        }
    }

    return {
        $set: set
    };
}

const hideColumn = function() {
    return visibility(...arguments, false);
}

const showColumn = function() {
    return visibility(...arguments, true);
}

const insertColumn = function(obj) {
    const [columns] = obj.args;

    columns.sort((a, b) => a.column - b.column);

    const worksheet = obj.document.spreadsheet.worksheets[obj.worksheetIndex];

    const currentData = worksheet.data;

    const isMatrix = dataIsMatrix(currentData);

    const currentColumns = worksheet.columns;

    const columnsLength = columns.length;

    for (let index = 0; index < columnsLength; index++) {
        const column = columns[index];

        const columnIndex = column.column;

        if (columnIndex > currentColumns.length) {
            currentColumns[columnIndex] = column.options || {};
        } else {
            currentColumns.splice(columnIndex, 0, column.options || {});
        }

        const columnData = column.data;

        const columnName = (column.options && column.options.name) || columnIndex;

        const numOfRows = Math.max(currentData.length, worksheet.minDimensions[1]);

        for (let rowIndex = 0; rowIndex < numOfRows; rowIndex++) {
            if (isMatrix === false) {
                if (!columnData || typeof columnData[rowIndex] === 'undefined' || columnData[rowIndex] === '') {
                    continue;
                }

                if (!currentData[rowIndex]) {
                    currentData[rowIndex] = {};
                }

                currentData[rowIndex][columnName] = columnData[rowIndex];
            } else {
                let cellValue = columnData && columnData[rowIndex];

                if (!currentData[rowIndex]) {
                    if (typeof cellValue === 'undefined' || cellValue === '') {
                        continue;
                    }

                    currentData[rowIndex] = [];
                }

                const row = currentData[rowIndex];

                if (columnIndex < row.length) {
                    row.splice(columnIndex, 0, cellValue);
                } else {
                    row[columnIndex] = cellValue;
                }
            }
        }
    }

    const worksheetPath = `spreadsheet.worksheets.${obj.worksheetIndex}`;

    const set = {
        [`${worksheetPath}.minDimensions.0`]: Math.max(worksheet.minDimensions[0] + columnsLength, columns[columnsLength - 1].column + 1),
        [`${worksheetPath}.data`]: currentData,
        [`${worksheetPath}.columns`]: currentColumns,
    };

    const changes = [{
        $set: set
    }];

    // Update references from other properties
    updateReferences(obj, changes);

    return changes;
}

const moveColumn = function (obj) {
    const [from, to, quantity] = obj.args;

    const worksheet = obj.document.spreadsheet.worksheets[obj.worksheetIndex];
    const columns = obj.document.spreadsheet.worksheets[obj.worksheetIndex].columns;

    moveArrayItems(columns, from, to, quantity);

    const worksheetPath = `spreadsheet.worksheets.${obj.worksheetIndex}`;

    const set = {
        [`${worksheetPath}.columns`]: columns,
    };

    if (worksheet.footers) {
        const footers = worksheet.footers;

        const footersLength = footers.length;
        for (let y = 0; y < footersLength; y++) {
            const footerRow = footers[y];

            moveArrayItems(footerRow, from, to, quantity);
        }

        set[`${worksheetPath}.footers`] = footers;
    }

    const data = worksheet.data;
    if (dataIsMatrix(data)) {
        const currentDataLength = data.length;
        for (let y = 0; y < currentDataLength; y++) {
            const row = data[y];

            if (row) {
                moveArrayItems(row, from, to, quantity);
            }
        }

        set[`${worksheetPath}.data`] = data;
    }

    const changes = [{
        $set: set,
    }];

    // Update references from other properties
    updateReferences(obj, changes);

    return changes;
}

const deleteColumn = function (obj) {
    const [columnsToDelete] = obj.args;

    columnsToDelete.sort((a, b) => a - b);

    let lastSimplifieRemoval = [columnsToDelete[0], 1];

    const simplifiedColumnsToDelete = [];

    const length = columnsToDelete.length;
    for (let i = 1; i < length; i++) {
        const columnIndex = columnsToDelete[i];

        if (columnIndex === lastSimplifieRemoval[0] + lastSimplifieRemoval[1]) {
            lastSimplifieRemoval[1]++;
        } else {
            simplifiedColumnsToDelete.push(lastSimplifieRemoval);

            lastSimplifieRemoval = [columnIndex, 1];
        }
    }

    simplifiedColumnsToDelete.push(lastSimplifieRemoval);

    const worksheet = obj.document.spreadsheet.worksheets[obj.worksheetIndex];

    const currentData = worksheet.data;
    const currentDataLength = currentData.length;

    const currentColumns = worksheet.columns;

    const isMatrix = dataIsMatrix(currentData);

    const footers = worksheet.footers;
    const footersLength = footers ? footers.length : 0;

    const worksheetPath = `spreadsheet.worksheets.${obj.worksheetIndex}`;

    const unset = {};

    for (let i = simplifiedColumnsToDelete.length - 1; i > -1; i--) {
        const [removalIndex, numOfRemovedItems] = simplifiedColumnsToDelete[i];

        const removedColumns = currentColumns.splice(removalIndex, numOfRemovedItems);

        if (footers) {
            for (let y = 0; y < footersLength; y++) {
                const footerRow = footers[y];

                footerRow.splice(removalIndex, numOfRemovedItems);
            }
        }

        for (let y = 0; y < currentDataLength; y++) {
            const row = currentData[y];

            if (isMatrix) {
                if (row) {
                    row.splice(removalIndex, numOfRemovedItems);
                }
            } else if (isMatrix === false) {
                const removedColumnsLength = removedColumns.length;
                for (let removedColumnIndex = 0; removedColumnIndex < removedColumnsLength; removedColumnIndex++) {
                    const removedColumn = removedColumns[removedColumnIndex];

                    if (removedColumn && removedColumn.name) {
                        unset[`${worksheetPath}.data.y.${removedColumn}`] = true;
                    }
                }
            }
        }
    }

    const set = {
        [`${worksheetPath}.columns`]: currentColumns,
        [`${worksheetPath}.minDimensions.0`]: worksheet.minDimensions[0] - length,
    };

    if (footers) {
        set[`${worksheetPath}.footers`] = footers;
    }

    let change = {
        $set: set,
    };

    if (isMatrix) {
        set[`${worksheetPath}.data`] = currentData;
    } else if (isMatrix === false) {
        change['$unset'] = unset;
    }

    const changes = [change];

    // Update references from other properties
    updateReferences(obj, changes);

    return changes;
}

module.exports = {
    setWidth,
    hideColumn,
    showColumn,
    insertColumn,
    moveColumn,
    deleteColumn,
}


/***/ },

/***/ 998
(module) {

const setComments = function(obj) {
    let comments = obj.args[0];

    let changes = {
        $set: {}
    };

    Object.keys(comments).forEach(key => {
        changes.$set[`spreadsheet.worksheets.${obj.worksheetIndex}.comments.${key}`] = comments[key];
    });

    return changes;
}

module.exports = {
    setComments,
}


/***/ },

/***/ 194
(module) {

const setConfig = function(obj) {
    let config = obj.args[0];
    let scope = obj.args[1];
    config = JSON.parse(config);
    let changes = {$set: {}};

    if (scope) {
        Object.keys(config).forEach((prop) => {
            changes.$set[`spreadsheet.${prop}`] = config[prop];
        });
    } else {
        Object.keys(config).forEach((prop) => {
            changes.$set[`spreadsheet.worksheets.${obj.worksheetIndex}.${prop}`] = config[prop];
        });
    }

    return changes;
}

module.exports = {
    setConfig,
}


/***/ },

/***/ 736
(module, __unused_webpack_exports, __webpack_require__) {

const { dataIsMatrix } = __webpack_require__(191);

const setValue = function(obj) {
    const valueChanges = obj.args[0];

    const worksheet = obj.document.spreadsheet.worksheets[obj.worksheetIndex];

    const currentData = worksheet.data;
    const columns = worksheet.columns;

    const isMatrix = dataIsMatrix(currentData);

    const dataPath = `spreadsheet.worksheets.${obj.worksheetIndex}.data`;

    const accessedRows = new Set();

    const insertRows = {};
    const setCellValues = {};

    const valueChangesLength = valueChanges.length;
    for (let i = 0; i < valueChangesLength; i++) {
        const { y, x, value } = valueChanges[i];

        const rowPath = `${dataPath}.${y}`;

        if (!accessedRows.has(y)) {
            if (!currentData[y]) {
                insertRows[rowPath] = isMatrix || typeof isMatrix !== 'boolean' ? [] : {};
            }

            accessedRows.add(y);
        }

        const column = isMatrix || typeof isMatrix !== 'boolean' ? x : (columns[x] && columns[x].name || x);

        setCellValues[`${rowPath}.${column}`] = value;
    }

    const changes = [];

    if (Object.keys(insertRows).length !== 0) {
        changes.push({
            $set: insertRows,
        });
    }

    changes.push({
        $set: setCellValues,
    });

    return changes;
}

const setFormula = setValue;

module.exports = {
    setValue,
    setFormula,
}


/***/ },

/***/ 227
(module) {

const setDefinedNames = function(obj) {
    let [definedNameChanges] = obj.args;

    const set = {};
    const unset = {};

    const definedNameChangesLength = definedNameChanges.length;
    for (let i = 0; i < definedNameChangesLength; i++) {
        const definedNameChange = definedNameChanges[i];

        const propertyPath = `spreadsheet.definedNames.${definedNameChange.index}`;

        if (typeof definedNameChange.value !== 'undefined') {
            set[propertyPath] = definedNameChange.value;
        } else {
            unset[propertyPath] = '';
        }
    }

    return {
        $set: set,
        $unset: unset,
    };
}

module.exports = {
    setDefinedNames,
}


/***/ },

/***/ 295
(module) {

const setFooter = function(obj) {
    let matrix = obj.args[0];

    let changes = { $set: {} };

    changes.$set[`spreadsheet.worksheets.${obj.worksheetIndex}.footers`] = matrix;

    return changes;
}

const setFooterValue = function(obj) {
    let items = obj.args[0]

    if (!Array.isArray(items)) {
        items = [items]
    }

    const numOfItems = items.length

    const set = {}

    for (let i = 0; i < numOfItems; i++) {
        const { x, y, value } = items[i];
        set[`spreadsheet.worksheets.${obj.worksheetIndex}.footers.${y}.${x}`] = value;
    }

    const changes = { $set: set };

    return changes;
}

const resetFooter = function(obj) {
    let changes = { $unset: {} };

    changes.$unset[`spreadsheet.worksheets.${obj.worksheetIndex}.footers`] = '';

    return changes;
}

module.exports = {
    setFooter,
    setFooterValue,
    resetFooter,
}


/***/ },

/***/ 348
(module) {

const setFreezeColumns = function(obj) {
    let columns = obj.args[0];

    let changes = { $set: {} };

    changes.$set[`spreadsheet.worksheets.${obj.worksheetIndex}.freezeColumns`] = columns;

    return changes;
}

module.exports = {
    setFreezeColumns,
}


/***/ },

/***/ 530
(module) {

const setFreezeRows = function(obj) {
    let rows = obj.args[0];

    let changes = { $set: {} };

    changes.$set[`spreadsheet.worksheets.${obj.worksheetIndex}.freezeRows`] = rows;

    return changes;
}

module.exports = {
    setFreezeRows,
}


/***/ },

/***/ 485
(module) {

const setGroup = function(obj, property) {
    const [index, quantity, state] = obj.args;

    const propertyPath = `spreadsheet.worksheets.${obj.worksheetIndex}.${property}.${index}`;

    if (quantity) {
        let set;

        const savedArray = obj.document.spreadsheet.worksheets[obj.worksheetIndex][property];

        if (savedArray[index]) {
            set = {
                [`${propertyPath}.group`]: quantity,
                [`${propertyPath}.state`]: state,
            }
        } else {
            set = {
                [propertyPath]: {
                    group: quantity,
                    state: state,
                }
            }
        }

        return {
            $set: set,
        };
    }

    return {
        $unset: {
            [`${propertyPath}.group`]: '',
            [`${propertyPath}.state`]: '',
        },
    };
}

const setColumnGroup = function(obj) {
    return setGroup(obj, 'columns');
}

const setRowGroup = function(obj) {
    return setGroup(obj, 'rows');
}

module.exports = {
    setColumnGroup,
    setRowGroup,
};

/***/ },

/***/ 701
(module) {

const setHeader = function(obj) {
    const [columnIndex, value] = obj.args;

    const columnPath = `spreadsheet.worksheets.${obj.worksheetIndex}.columns.${columnIndex}`;

    const savedColumns = obj.document.spreadsheet.worksheets[obj.worksheetIndex].columns;

    const set = {};

    if (savedColumns[columnIndex]) {
        set[`${columnPath}.title`] = value;
    } else {
        set[columnPath] = {
            title: value,
        };
    }

    return {
        $set: set,
    };
}

module.exports = {
    setHeader,
}


/***/ },

/***/ 502
(module) {

const setMedia = function(obj) {
    let mediaChanges = obj.args[0];

    if (!Array.isArray(mediaChanges)) {
        mediaChanges = [mediaChanges];
    }

    const set = {};
    const pull = [];
    const push = [];

    let media = obj.document.spreadsheet.worksheets[obj.worksheetIndex].media;

    const mediaPath = `spreadsheet.worksheets.${obj.worksheetIndex}.media`;

    const mediaChangesLength = mediaChanges.length;
    for (let i = 0; i < mediaChangesLength; i++) {
        const item = mediaChanges[i];

        if (Object.keys(item).length !== 1) {
            const mediaIndex = media.findIndex((mediaItem) => mediaItem.id === item.id);

            if (mediaIndex > -1) {
                const entries = Object.entries(item);

                const mediaIndexPath = `${mediaPath}.${mediaIndex}`;

                const entriesLength = entries.length;
                for (let entryIndex = 0; entryIndex < entriesLength; entryIndex++) {
                    const entry = entries[entryIndex];

                    set[`${mediaIndexPath}.${entry[0]}`] = entry[1];
                }
            } else {
                push.push(item);
            }
        } else {
            pull.push(item.id);
        }
    }

    const result = {};

    if (Object.keys(set).length !== 0) {
        result['$set'] = set;
    }

    if (push.length !== 0) {
        result['$push'] = {
            [mediaPath]: {
                $each: push,
            },
        };
    }

    if (pull.length !== 0) {
        result['$pull'] = {
            [mediaPath]: {
                id: {
                    $in: pull,
                }
            },
        };
    }

    return result;
}

module.exports = {
    setMedia,
}


/***/ },

/***/ 562
(module) {

const setMerge = function(obj) {
    let mergeObject = obj.args[0];

    let changes = { $set: {} };

    Object.keys(mergeObject).forEach((key) => {
        changes.$set[`spreadsheet.worksheets.${obj.worksheetIndex}.mergeCells.${key}`] = mergeObject[key];
    });

    return changes;
}

const removeMerge = function(obj) {
    let mergeObject = obj.args[0];

    const cellNames = Object.keys(mergeObject);
    const cellNamesLength = cellNames.length;

    const unset = {};

    const propertyPath = `spreadsheet.worksheets.${obj.worksheetIndex}.mergeCells`;

    for (let i = 0; i < cellNamesLength; i++) {
        const cellName = cellNames[i];

        unset[`${propertyPath}.${cellName}`] = '';
    }

    return {
        $unset: unset,
    };
}

module.exports = {
    setMerge,
    removeMerge,
}


/***/ },

/***/ 657
(module) {

const setMeta = function(obj) {
    let metaObject = obj.args[0];

    let changes = { $set: {} };

    Object.keys(metaObject).forEach(key => {
        changes.$set[`spreadsheet.worksheets.${obj.worksheetIndex}.meta.${key}`] = metaObject[key];
    });

    return changes;
}

const resetMeta = function(obj) {
    let cellNames = obj.args[0];

    let changes = { $unset: {} };

    if (cellNames) {
        cellNames.forEach((cellName) => {
            changes.$unset[`spreadsheet.worksheets.${obj.worksheetIndex}.meta.${cellName}`] = '';
        });
    } else {
        changes.$unset[`spreadsheet.worksheets.${obj.worksheetIndex}.meta`] = '';
    }

    return changes;
}

module.exports = {
    setMeta,
    resetMeta,
}


/***/ },

/***/ 349
(module) {

const setNestedHeaders = function(obj) {
    let matrix = obj.args[0];

    let changes = { $set: {} };

    changes.$set[`spreadsheet.worksheets.${obj.worksheetIndex}.nestedHeaders`] = matrix;

    return changes;
}

const setNestedCell = function(obj) {
    const [cellChanges] = obj.args;

    const set = {};

    const nestedHeadersPath = `spreadsheet.worksheets.${obj.worksheetIndex}.nestedHeaders`;

    const cellChangesLength = cellChanges.length;
    for (let i = 0; i < cellChangesLength; i++) {
        const cellChange = cellChanges[i];

        const entries = Object.entries(cellChange.properties);

        const cellPath = `${nestedHeadersPath}.${cellChange.y}.${cellChange.x}`;

        const entriesLength = entries.length;
        for (let entryIndex = 0; entryIndex < entriesLength; entryIndex++) {
            const [key, value] = entries[entryIndex];

            set[`${cellPath}.${key}`] = value;
        }
    }

    return {
        $set: set,
    };
}

const resetNestedHeaders = function(obj) {
    
    let changes = { $unset: {} };

    changes.$unset[`spreadsheet.worksheets.${obj.worksheetIndex}.nestedHeaders`] = '';

    return changes;
}

module.exports = {
    setNestedHeaders,
    setNestedCell,
    resetNestedHeaders,
}


/***/ },

/***/ 574
(module, __unused_webpack_exports, __webpack_require__) {

const { updateReferences } = __webpack_require__(191);

const orderBy = function(obj) {
    const newOrder = obj.args[2];

    const worksheet = obj.document.spreadsheet.worksheets[obj.worksheetIndex];

    const currentData = worksheet.data;
    const currentRows = worksheet.rows;

    const newData = [];
    const newRows = [];

    const newOrderLength = newOrder.length;
    for (let i = 0; i < newOrderLength; i++) {
        const oldIndex = newOrder[i];

        newData.push(currentData[oldIndex]);
        newRows.push(currentRows[oldIndex]);
    }

    const worksheetPath = `spreadsheet.worksheets.${obj.worksheetIndex}`;

    const changes = [{
        $set: {
            [`${worksheetPath}.data`]: newData,
            [`${worksheetPath}.rows`]: newRows,
        }
    }];

    // Update references from other properties
    updateReferences(obj, changes);

    return changes;
}

module.exports = {
    orderBy,
}


/***/ },

/***/ 381
(module, __unused_webpack_exports, __webpack_require__) {

const { getColumnNameFromCoords } = __webpack_require__(582);

const setProperty = function(obj) {
    let propertyChanges = obj.args[0];

    const set = {};

    const worksheetPath = `spreadsheet.worksheets.${obj.worksheetIndex}`;

    const columnsPath = `${worksheetPath}.columns`;
    const cellsPath = `${worksheetPath}.cells`;

    const propertyChangesLength = propertyChanges.length;
    for (let i = 0; i < propertyChangesLength; i++) {
        let propertyPath;
        let propertyChange = propertyChanges[i];

        if (typeof propertyChange.y === 'undefined' || propertyChange.y === null) {
            // Get the column path
            propertyPath = `${columnsPath}.${propertyChange.x}`;
        } else {
            // Get the cell name
            let cellName = getColumnNameFromCoords(propertyChange.x, propertyChange.y);
            // Get the cell path
            propertyPath = `${cellsPath}.${cellName}`;
        }

        // Update to the following value
        set[propertyPath] = typeof propertyChange.value !== 'undefined' ? propertyChange.value : null;
    }

    return {
        $set: set,
    };
}

const updateProperty = function(obj) {
    let propertyChanges = obj.args[0];

    const worksheetPath = `spreadsheet.worksheets.${obj.worksheetIndex}`;

    const cellsPath = `${worksheetPath}.cells`;
    const columnsPath = `${worksheetPath}.columns`;

    const columns = obj.document.spreadsheet.worksheets[obj.worksheetIndex].columns;

    const setNewColumns = {};
    const set = {};

    const propertyChangesLength = propertyChanges.length;
    for (let i = 0; i < propertyChangesLength; i++) {
        const propertyChange = propertyChanges[i];

        const entries = Object.entries(propertyChange.value);

        if (typeof propertyChange.y === 'undefined' || propertyChange.y === null) {
            const columnPath = `${columnsPath}.${propertyChange.x}`;

            if (!columns[propertyChange.x]) {
                setNewColumns[columnPath] = propertyChange.value;

                columns[propertyChange.x] = {};
            } else {
                const entriesLength = entries.length;
                for (let entryIndex = 0; entryIndex < entriesLength; entryIndex++) {
                    const [key, value] = entries[entryIndex];

                    set[`${columnPath}.${key}`] = value;
                }
            }
        } else {
            const cellPath = `${cellsPath}.${getColumnNameFromCoords(propertyChange.x, propertyChange.y)}`;

            const entriesLength = entries.length;
            for (let entryIndex = 0; entryIndex < entriesLength; entryIndex++) {
                const [key, value] = entries[entryIndex];

                set[`${cellPath}.${key}`] = value;
            }
        }
    }

    const changes = [];

    if (Object.keys(setNewColumns).length !== 0) {
        changes.push({
            $set: setNewColumns
        });
    }

    if (Object.keys(set).length !== 0) {
        changes.push({
            $set: set
        });
    }

    return changes;
}

module.exports = {
    setProperty,
    updateProperty,
}


/***/ },

/***/ 954
(module, __unused_webpack_exports, __webpack_require__) {

const { updateReferences, dataIsMatrix, moveArrayItems } = __webpack_require__(191);

const setHeight = function (obj) {
    let [rows, heights] = obj.args;

    if (!Array.isArray(rows)) {
        rows = [rows];
    }

    const set = {};

    const heightsIsAnArray = Array.isArray(heights);
    const rowsPath = `spreadsheet.worksheets.${obj.worksheetIndex}.rows`;
    const savedRows = obj.document.spreadsheet.worksheets[obj.worksheetIndex].rows || [];

    const rowsLength = rows.length;

    for (let i = 0; i < rowsLength; i++) {
        const rowIndex = rows[i];
        const rowPath = `${rowsPath}.${rowIndex}`;
        const height = heightsIsAnArray ? heights[i] : heights;

        if (savedRows[rowIndex]) {
            set[`${rowPath}.height`] = height;
        } else {
            set[rowPath] = { height: height };
        }
    }

    return {
        $set: set,
    };
}

const visibility = function(obj, state) {
    let rows = obj.args[0];

    if (!Array.isArray(rows)) {
        rows = [rows];
    }

    const set = {};

    const rowsPath = `spreadsheet.worksheets.${obj.worksheetIndex}.rows`;
    const savedRows = obj.document.spreadsheet.worksheets[obj.worksheetIndex].rows || [];

    const rowsLength = rows.length;

    for (let i = 0; i < rowsLength; i++) {
        const rowIndex = rows[i];

        const rowPath = `${rowsPath}.${rowIndex}`;

        if (savedRows[rowIndex]) {
            set[`${rowPath}.visible`] = state;
        } else {
            set[rowPath] = {
                visible: state,
            };
        }
    }

    return {
        $set: set,
    };
}

const turnRowIntoObj = function(row, columns) {
    const result = {};

    const rowLength = row.length;
    for (let columnIndex = 0; columnIndex < rowLength; columnIndex++) {
        const propertyName = (columns[columnIndex] && columns[columnIndex].name) || columnIndex;

        result[propertyName] = row[columnIndex];
    }

    return result;
}

const insertRow = function(obj) {
    const [rows] = obj.args;

    rows.sort((a, b) => a.row - b.row);

    const worksheet = obj.document.spreadsheet.worksheets[obj.worksheetIndex];

    const currentData = worksheet.data;
    let currentDataLength = currentData.length;

    const isMatrix = dataIsMatrix(currentData);

    const worksheetPath = `spreadsheet.worksheets.${obj.worksheetIndex}`;

    const currentColumns = worksheet.columns;

    const currentRows = worksheet.rows;

    const rowsLength = rows.length;
    for (let index = 0; index < rowsLength; index++) {
        const { row: rowIndex, data: rowData, options } = rows[index];

        const data = isMatrix === false && Array.isArray(rowData) ? turnRowIntoObj(rowData, currentColumns) : rowData;

        if (rowIndex > currentDataLength) {
            currentData[rowIndex] = data;

            currentDataLength = rowIndex + 1;
        } else {
            currentData.splice(rowIndex, 0, data);

            currentDataLength++;
        }

        if (rowIndex > currentRows.length) {
            currentRows[rowIndex] = options;
        } else {
            currentRows.splice(rowIndex, 0, options);
        }
    }

    const set = {
        [`${worksheetPath}.data`]: currentData,
        [`${worksheetPath}.rows`]: currentRows,
        [worksheetPath + '.minDimensions.1']: Math.max(worksheet.minDimensions[1] + rowsLength, rows[rowsLength - 1].row + 1),
    };

    const changes = [{
        $set: set
    }];

    // Update references from other properties
    updateReferences(obj, changes);

    return changes;
}

const hideRow = function (obj) {
    return visibility(obj, false)
}

const showRow = function (obj) {
    return visibility(obj, true)
}

const moveRow = function (obj) {
    const [from, to, quantity] = obj.args;

    const worksheet = obj.document.spreadsheet.worksheets[obj.worksheetIndex];

    const rows = worksheet.rows;

    moveArrayItems(rows, from, to, quantity);

    const data = worksheet.data;

    moveArrayItems(data, from, to, quantity);

    const worksheetPath = `spreadsheet.worksheets.${obj.worksheetIndex}`;

    const changes = [{
        $set: {
            [`${worksheetPath}.data`]: data,
            [`${worksheetPath}.rows`]: rows
        }
    }];

    // Update references from other properties
    updateReferences(obj, changes);

    return changes;
}

const deleteRow = function(obj) {
    const [rowsToDelete] = obj.args;

    rowsToDelete.sort((a, b) => a - b);

    let lastSimplifieRemoval = [rowsToDelete[0], 1];

    const simplifiedRowsToDelete = [];

    const length = rowsToDelete.length;
    for (let i = 1; i < length; i++) {
        const rowIndex = rowsToDelete[i];

        if (rowIndex === lastSimplifieRemoval[0] + lastSimplifieRemoval[1]) {
            lastSimplifieRemoval[1]++;
        } else {
            simplifiedRowsToDelete.push(lastSimplifieRemoval);

            lastSimplifieRemoval = [rowIndex, 1];
        }
    }

    simplifiedRowsToDelete.push(lastSimplifieRemoval);

    const worksheet = obj.document.spreadsheet.worksheets[obj.worksheetIndex];

    const currentData = worksheet.data;
    const currentRows = worksheet.rows;

    for (let i = simplifiedRowsToDelete.length - 1; i > -1; i--) {
        const [removalIndex, numOfRemovedItems] = simplifiedRowsToDelete[i];

        currentData.splice(removalIndex, numOfRemovedItems);
        currentRows.splice(removalIndex, numOfRemovedItems);
    }

    const worksheetPath = `spreadsheet.worksheets.${obj.worksheetIndex}`;

    const set = {
        [`${worksheetPath}.data`]: currentData,
        [`${worksheetPath}.rows`]: currentRows,
        [`${worksheetPath}.minDimensions.1`]: worksheet.minDimensions[1] - length
    }

    const changes = [{
        $set: set,
    }];

    // Update references from other properties
    updateReferences(obj, changes);

    return changes;
}

const setRowId = function(obj) {
    const savedRows = obj.document.spreadsheet.worksheets[obj.worksheetIndex].rows;

    const rowsPath = `spreadsheet.worksheets.${obj.worksheetIndex}.rows`;

    const set = {};

    const idChanges = Object.entries(obj.args[0]);

    const idChangesLength = idChanges.length;
    for (let i = 0; i < idChangesLength; i++) {
        const [rowIndex, rowId] = idChanges[i];

        const rowPath = `${rowsPath}.${rowIndex}`;

        if (savedRows[rowIndex]) {
            set[`${rowPath}.id`] = rowId;
        } else {
            set[rowPath] = {
                id: rowId
            };
        }
    }

    return {
        $set: set,
    };
}

module.exports = {
    setHeight,
    insertRow,
    hideRow,
    showRow,
    moveRow,
    deleteRow,
    setRowId,
}


/***/ },

/***/ 315
(module) {

const setStyle = function(obj) {
    const set = {
        [`spreadsheet.style`]: obj.instance.config.style,
        [`spreadsheet.worksheets.${obj.worksheetIndex}.style`]: obj.instance.worksheets[obj.worksheetIndex].getStyle(null, true),
    };

    return {
        $set: set,
    };
}

const resetStyle = function(obj) {
    const set = {
        [`spreadsheet.worksheets.${obj.worksheetIndex}.style`]: obj.instance.worksheets[obj.worksheetIndex].getStyle(null, true),
    };

    return {
        $set: set,
    };
}

module.exports = {
    setStyle,
    resetStyle
}


/***/ },

/***/ 774
(module) {

const setValidations = function(obj) {
    let newValue = obj.args[0];

    const set = {}

    newValue.forEach((validation) => {
        let key = `spreadsheet.validations.${validation.index}`;
        if (typeof validation.value === 'string') {
            key += '.range';
        }

        set[key] = validation.value;
    })

    const changes = { $set: set };

    return changes;
}

module.exports = {
    setValidations,
}


/***/ },

/***/ 666
(module) {

const validateWorksheet = function(worksheet) {
    // Media
    if (! worksheet.media) {
        worksheet.media = [];
    }
    // Data
    if (! worksheet.data) {
        worksheet.data = [];
    }
    // Rows
    if (! worksheet.rows) {
        worksheet.rows = [];
    }
    // Columns
    if (! worksheet.columns) {
        worksheet.columns = [];
    }
}

const createWorksheet = function(obj) {
    let config = obj.args[0];
    let position = obj.args[1]

    validateWorksheet(config);

    let changes = {
        $push: {
            'spreadsheet.worksheets': {
                $each: [config],
            }
        }
    };

    if (typeof(position) !== 'undefined') {
        changes.$push['spreadsheet.worksheets'].$position = position;
    }

    return changes;
}

const deleteWorksheet = function(obj) {
    let position = obj.args[0]

    let unset = {}
    unset[`spreadsheet.worksheets.${position}`] = 1;

    let pull = {};
    pull[`spreadsheet.worksheets`] = null;

    return [{ $unset: unset }, { $pull: pull }];
}

const renameWorksheet = function(obj) {
    let worksheetIndex = obj.args[0];
    let newName = obj.args[1]
    let changes = { $set: {} };
    changes.$set[`spreadsheet.worksheets.${worksheetIndex}.worksheetName`] = newName;

    return changes;
}

const moveWorksheet = function(obj) {
    let origin = obj.args[0];
    let destination = obj.args[1];

    let worksheet = obj.document.spreadsheet.worksheets[origin];
    if (! worksheet) {
        throw new Error('Worksheet not found');
    }

    const pull = {};
    pull['spreadsheet.worksheets'] = worksheet;

    let push = {};
    push['spreadsheet.worksheets'] = { $each: [worksheet], $position: destination };

    return [{ $pull: pull },{ $push: push }];
}

const setWorksheetState = function(obj) {
    let [index, state] = obj.args;
    let changes = { $set: {} };

    if (typeof state === 'boolean') {
        state = state ? 'visible' : 'hidden';
    }

    changes.$set[`spreadsheet.worksheets.${index}.worksheetState`] = state;
    return changes;
}

module.exports = {
    validateWorksheet,
    createWorksheet,
    deleteWorksheet,
    renameWorksheet,
    moveWorksheet,
    setWorksheetState,
}


/***/ },

/***/ 191
(module) {

const updateReferences = function(obj, changes) {
    const set = {};

    // Properties to be updated
    let properties = ['style','meta','comments','cells','mergeCells'];
    properties.forEach((v) => {
        let values = obj.instance.worksheets[obj.worksheetIndex].options[v];
        if (values && Object.keys(values).length > 0) {
            set[`spreadsheet.worksheets.${obj.worksheetIndex}.${v}`] = values;
        }
    });
    changes.push({ $set: set })
}

const moveArrayItems = function(array, from, to, quantity) {
    const movedItems = array.splice(from, quantity);

    while (movedItems.length < quantity) {
        movedItems.push(null);
    }

    const insertAt = from < to ? to - quantity + 1 : to;

    while (array.length < insertAt) {
        array.push(null);
    }

    array.splice(insertAt, 0, ...movedItems);
}

const dataIsMatrix = function(data) {
    let dataLength = data.length;
    for (let i = 0; i < dataLength; i++) {
        if (data[i]) {
            return Array.isArray(data[i]);
        }
    }

    return null;
}

module.exports = {
    updateReferences,
    moveArrayItems,
    dataIsMatrix,
}

/***/ },

/***/ 127
(module) {

"use strict";
module.exports = require("jsonwebtoken");

/***/ },

/***/ 582
(module) {

"use strict";
module.exports = require("jspreadsheet");

/***/ },

/***/ 884
(module) {

"use strict";
module.exports = require("mongodb");

/***/ },

/***/ 333
(module) {

"use strict";
module.exports = require("uuid");

/***/ }

/******/ 	});
/************************************************************************/
/******/ 	// The module cache
/******/ 	var __webpack_module_cache__ = {};
/******/ 	
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/ 		// Check if module is in cache
/******/ 		var cachedModule = __webpack_module_cache__[moduleId];
/******/ 		if (cachedModule !== undefined) {
/******/ 			return cachedModule.exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = __webpack_module_cache__[moduleId] = {
/******/ 			// no module.id needed
/******/ 			// no module.loaded needed
/******/ 			exports: {}
/******/ 		};
/******/ 	
/******/ 		// Execute the module function
/******/ 		__webpack_modules__[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/ 	
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/ 	
/************************************************************************/
/******/ 	
/******/ 	// startup
/******/ 	// Load entry module and return exports
/******/ 	// This entry module is referenced by other modules so it can't be inlined
/******/ 	var __webpack_exports__ = __webpack_require__(44);
/******/ 	module.exports = __webpack_exports__;
/******/ 	
/******/ })()
;
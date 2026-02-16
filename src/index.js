const uuid = require("uuid");
const jwt = require('jsonwebtoken');
const { MongoClient } = require('mongodb');

const { setWidth, hideColumn, showColumn, insertColumn, moveColumn, deleteColumn } = require("./modules/column");
const { setHeight, hideRow, showRow, moveRow, insertRow, deleteRow, setRowId } = require("./modules/row");
const { setMeta, resetMeta } = require("./modules/meta");
const { setComments } = require("./modules/comments");
const { setFreezeColumns } = require("./modules/freezeColumns");
const { setFreezeRows } = require("./modules/freezeRows");
const { setFooter, setFooterValue, resetFooter } = require("./modules/footer");
const { setHeader } = require("./modules/header");
const { setNestedHeaders, resetNestedHeaders, setNestedCell } = require("./modules/nestedHeaders");
const { setMerge, removeMerge } = require("./modules/merge");
const { setCache } = require("./modules/cache");
const { setStyle, resetStyle } = require("./modules/style");
const { setProperty, updateProperty } = require("./modules/property");
const { setValidations } = require("./modules/validations");
const { setMedia } = require("./modules/media");
const { createWorksheet, deleteWorksheet, renameWorksheet, moveWorksheet,setWorksheetState, validateWorksheet } = require("./modules/worksheet");
const { setValue, setFormula } = require("./modules/data");
const { setColumnGroup, setRowGroup } = require("./modules/group");
const { setDefinedNames } = require("./modules/definedNames");
const { setConfig } = require("./modules/config");
const { orderBy } = require('./modules/order');

;(function(global, factory) {
    typeof exports === 'object' && typeof module !== 'undefined' ? module.exports = factory() :
    typeof define === 'function' && define.amd ? define(factory) :
    global.adapter = factory();
}(this, (function() {

    const url = process.env.MONGODB_URI || process.env.MONGODB_URL || 'mongodb://mongodb';
    const dbName = process.env.MONGODB_DB_NAME || 'jspreadsheet';
    const collectionName = process.env.MONGODB_COLLECTION_NAME || 'documents';

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

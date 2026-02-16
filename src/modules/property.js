const { getColumnNameFromCoords } = require("jspreadsheet");

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

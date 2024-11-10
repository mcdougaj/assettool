let excelData = null;
let columnMappings = {
    parent: null,
    child: null,
    description: null
};
let selectedNode = null;
let originalWorkbook = null;
let lastTreeData = null;
let historyStack = []; // Add history stack
let editHistory = [];

document.addEventListener('DOMContentLoaded', function() {
    initializeFileLoader();
    initializeDragAndDrop();
    initializeTreeView();
    initializeEditPanel();
    initializeNewRecordModal();
    initializeSearch();
    initializeResizablePanels();
    initializeCollapsibleSections();
    initializeUndoButton();
    initializeOrphanSearch();
});

function initializeUndoButton() {
    document.getElementById('undoHierarchy').addEventListener('click', undoLastChange);
}

function undoLastChange() {
    if (historyStack.length > 0) {
        excelData = JSON.parse(historyStack.pop()); // Restore previous state
        updateTreeView();
    } else {
        alert('No more changes to undo');
    }
}

function initializeCollapsibleSections() {
    const sections = document.querySelectorAll('.collapsible-section');
    sections.forEach(section => {
        const header = section.querySelector('.section-header');
        const content = section.querySelector('.section-content');
        const toggleIcon = header.querySelector('.toggle-icon');

        header.addEventListener('click', function() {
            content.classList.toggle('collapsed');
            toggleIcon.classList.toggle('collapsed');
        });
    });
}

function initializeFileLoader() {
    document.getElementById('loadFile').addEventListener('click', function() {
        const fileInput = document.getElementById('excelFile');
        const file = fileInput.files[0];
        if (file) {
            const reader = new FileReader();
            reader.onload = function(e) {
                const data = new Uint8Array(e.target.result);
                originalWorkbook = XLSX.read(data, { type: 'array' });
                const firstSheet = originalWorkbook.Sheets[originalWorkbook.SheetNames[0]];
                excelData = XLSX.utils.sheet_to_json(firstSheet);
                populateColumnsList();
            };
            reader.readAsArrayBuffer(file);
        } else {
            alert('Please select a file first');
        }
    });
}

function initializeResizablePanels() {
    const leftPanel = document.getElementById('leftPanel');
    const rightPanel = document.getElementById('rightPanel');
    const leftHandle = document.querySelector('.left-handle');
    const rightHandle = document.querySelector('.right-handle');

    let isResizing = false;
    let currentHandle = null;

    const startResize = function(e, handle, panel) {
        isResizing = true;
        currentHandle = handle;
        document.body.style.cursor = 'col-resize';
        handle.classList.add('active');
        
        const startX = e.pageX;
        const startWidth = panel.offsetWidth;
        
        const handleMouseMove = function(e) {
            if (!isResizing) return;
            
            const diff = e.pageX - startX;
            let newWidth;
            
            if (handle.classList.contains('left-handle')) {
                newWidth = startWidth + diff;
            } else {
                newWidth = startWidth - diff;
            }
            
            newWidth = Math.max(250, Math.min(500, newWidth));
            panel.style.width = `${newWidth}px`;
        };
        
        const stopResize = function() {
            isResizing = false;
            document.body.style.cursor = '';
            handle.classList.remove('active');
            document.removeEventListener('mousemove', handleMouseMove);
            document.removeEventListener('mouseup', stopResize);
        };
        
        document.addEventListener('mousemove', handleMouseMove);
        document.addEventListener('mouseup', stopResize);
    };

    leftHandle.addEventListener('mousedown', (e) => startResize(e, leftHandle, leftPanel));
    rightHandle.addEventListener('mousedown', (e) => startResize(e, rightHandle, rightPanel));
}

function initializeNewRecordModal() {
    const modal = document.getElementById('newRecordModal');
    const addButton = document.getElementById('addNewRecord');
    const saveButton = document.getElementById('saveNewRecord');
    const cancelButton = document.getElementById('cancelNewRecord');

    addButton.addEventListener('click', function() {
        if (!columnMappings.parent || !columnMappings.child) {
            alert('Please map the Parent and Child columns first');
            return;
        }
        modal.classList.add('show');
    });

    saveButton.addEventListener('click', function() {
        const parentValue = document.getElementById('newParentValue').value.trim();
        const childValue = document.getElementById('newChildValue').value.trim();
        const description = document.getElementById('newDescription').value.trim();

        if (!parentValue || !childValue) {
            alert('Parent and Child values are required');
            return;
        }

        historyStack.push(JSON.stringify(excelData)); // Save current state
        const newRecord = {};
        newRecord[columnMappings.parent] = parentValue;
        newRecord[columnMappings.child] = childValue;
        if (columnMappings.description) {
            newRecord[columnMappings.description] = description;
        }

        excelData.push(newRecord);
        updateSummaries(); // Update summaries after adding
        updateTreeView();
        closeNewRecordModal();
        clearNewRecordForm();
    });

    cancelButton.addEventListener('click', closeNewRecordModal);
}

function closeNewRecordModal() {
    const modal = document.getElementById('newRecordModal');
    modal.classList.remove('show');
}

function clearNewRecordForm() {
    document.getElementById('newParentValue').value = '';
    document.getElementById('newChildValue').value = '';
    document.getElementById('newDescription').value = '';
}

function initializeSearch() {
    const searchButton = document.getElementById('searchButton');
    const clearButton = document.getElementById('clearSearch');
    const searchInput = document.getElementById('searchInput');

    searchButton.addEventListener('click', performSearch);
    clearButton.addEventListener('click', clearSearch);
    searchInput.addEventListener('keypress', function(e) {
        if (e.key === 'Enter') {
            performSearch();
        }
    });
}

function performSearch() {
    const searchTerm = document.getElementById('searchInput').value.toLowerCase();
    const searchParent = document.getElementById('searchParent').checked;
    const searchChild = document.getElementById('searchChild').checked;
    const searchDescription = document.getElementById('searchDescription').checked;

    if (!searchTerm || !excelData) return;

    const searchResults = excelData.filter(row => {
        let match = false;
        if (searchParent && row[columnMappings.parent]) {
            match = match || row[columnMappings.parent].toString().toLowerCase().includes(searchTerm);
        }
        if (searchChild && row[columnMappings.child]) {
            match = match || row[columnMappings.child].toString().toLowerCase().includes(searchTerm);
        }
        if (searchDescription && columnMappings.description && row[columnMappings.description]) {
            match = match || row[columnMappings.description].toString().toLowerCase().includes(searchTerm);
        }
        return match;
    });

    if (searchResults.length === 0) {
        alert('No matches found');
        return;
    }

    const treeData = buildTreeFromSearchResults(searchResults);
    $('#treeView').jstree(true).settings.core.data = treeData;
    $('#treeView').jstree(true).refresh();
}

function buildTreeFromSearchResults(results) {
    const treeData = [];
    const processedNodes = new Set();
    const parentChildMap = new Map();

    results.forEach(row => {
        const parentValue = row[columnMappings.parent];
        const childValue = row[columnMappings.child];
        const description = columnMappings.description ? row[columnMappings.description] : '';

        if (parentValue) {
            if (!parentChildMap.has(parentValue)) {
                parentChildMap.set(parentValue, []);
            }
            parentChildMap.get(parentValue).push({ child: childValue, description });
        }
    });

    function addNode(value, isParent = true, level = 0) {
        if (processedNodes.has(value)) return null;
        if (!value) return null;
        
        processedNodes.add(value);

        const icons = ['fas fa-folder', 'fas fa-folder-open', 'fas fa-toolbox'];
        const node = {
            text: value,
            id: value,
            children: [],
            state: { opened: true },
            icon: isParent ? icons[0] : icons[2]
        };

        if (parentChildMap.has(value)) {
            parentChildMap.get(value).forEach(({ child, description }) => {
                const childNode = addNode(child, false, level + 1);
                if (childNode) {
                    childNode.text = child + (description ? ` - ${description}` : '');
                    node.children.push(childNode);
                    node.icon = icons[1]; // Change to open folder if it has children
                }
            });
        }

        return node;
    }

    results.forEach(row => {
        const parentValue = row[columnMappings.parent];
        if (!processedNodes.has(parentValue)) {
            const node = addNode(parentValue);
            if (node) treeData.push(node);
        }
    });

    return treeData;
}

function clearSearch() {
    document.getElementById('searchInput').value = '';
    document.getElementById('orphanSearchContainer').style.display = 'none';
    window.currentOrphans = null; // Clear orphans data
    updateTreeView();
}

function populateColumnsList() {
    const columnsList = document.getElementById('columnsList');
    columnsList.innerHTML = '';
    
    if (excelData && excelData.length > 0) {
        const columns = Object.keys(excelData[0]);
        columns.forEach(column => {
            const columnDiv = document.createElement('div');
            columnDiv.className = 'column-item';
            columnDiv.draggable = true;
            columnDiv.textContent = column;
            columnDiv.dataset.column = column;
            
            columnDiv.addEventListener('dragstart', handleDragStart);
            columnsList.appendChild(columnDiv);
        });
    }
}

function initializeDragAndDrop() {
    const dropTargets = document.querySelectorAll('.drop-target');
    
    dropTargets.forEach(target => {
        target.addEventListener('dragover', handleDragOver);
        target.addEventListener('dragleave', handleDragLeave);
        target.addEventListener('drop', handleDrop);
    });
}

function handleDragStart(e) {
    e.dataTransfer.setData('text/plain', e.target.dataset.column);
}

function handleDragOver(e) {
    e.preventDefault();
    e.currentTarget.classList.add('dragover');
}

function handleDragLeave(e) {
    e.currentTarget.classList.remove('dragover');
}

function handleDrop(e) {
    e.preventDefault();
    e.currentTarget.classList.remove('dragover');
    const column = e.dataTransfer.getData('text/plain');
    e.currentTarget.textContent = column;
    
    historyStack.push(JSON.stringify(excelData)); // Save current state
    const mappingType = e.currentTarget.dataset.type;
    columnMappings[mappingType] = column;
    
    updateSummaries(); // Update summaries after mapping
    updateTreeView();
}

function initializeTreeView() {
    $('#treeView').jstree({
        core: {
            check_callback: true,
            data: [],
            themes: {
                name: 'default',
                dots: true,
                icons: true,
                variant: 'large'
            }
        },
        plugins: ['dnd', 'wholerow']
    }).on('select_node.jstree', function(e, data) {
        selectedNode = data.node;
        updateEditForm(selectedNode);
    }).on('move_node.jstree', function(e, data) {
        // Save current state before updating
        historyStack.push(JSON.stringify(excelData));
        
        const movedNodeId = data.node.id;
        const newParentId = data.parent === '#' ? '' : data.parent;
        
        // Find and update the record in excelData
        const record = excelData.find(row => 
            row[columnMappings.child] === movedNodeId ||
            row[columnMappings.parent] === movedNodeId
        );

        if (record) {
            // If the node was a child, update its parent
            if (record[columnMappings.child] === movedNodeId) {
                record[columnMappings.parent] = newParentId;
            }
            
            // Update the edit form if the moved node is selected
            if (selectedNode && selectedNode.id === movedNodeId) {
                updateEditForm(data.node);
            }
        }

        updateSummaries(); // Update summaries after moving
        updateTreeView(); // Refresh tree view
    });

    // Add expand/collapse handlers
    document.getElementById('expandAll').addEventListener('click', function() {
        $('#treeView').jstree('open_all');
    });
    
    document.getElementById('collapseAll').addEventListener('click', function() {
        $('#treeView').jstree('close_all');
    });

    document.getElementById('showOrphans').addEventListener('click', showOrphanRecords);
    document.getElementById('saveChanges').addEventListener('click', saveChanges);
}

function initializeEditPanel() {
    document.getElementById('applyChanges').addEventListener('click', applyChanges);
    document.getElementById('cancelChanges').addEventListener('click', cancelChanges);
    document.getElementById('undoEditChanges').addEventListener('click', undoEditChanges);
}

function undoEditChanges() {
    if (editHistory.length > 0) {
        const lastState = editHistory.pop();
        const rowIndex = excelData.findIndex(row => 
            row[columnMappings.parent] === lastState[columnMappings.parent] || 
            row[columnMappings.child] === lastState[columnMappings.child]
        );

        if (rowIndex !== -1) {
            excelData[rowIndex] = { ...lastState };
            updateTreeView();
            updateEditForm(selectedNode);
        }
    } else {
        alert('No more edit changes to undo');
    }
}

function updateEditForm(node) {
    const editForm = document.getElementById('editForm');
    const editActions = document.querySelector('.edit-actions');
    const editPanelTitle = document.getElementById('editPanelTitle');
    const editFormContainer = document.querySelector('.edit-form-container');
    
    if (!node || !excelData) {
        editHistory = [];
        editForm.innerHTML = '<p class="placeholder-text">Select a node in the tree to edit its values</p>';
        editActions.style.display = 'none';
        editPanelTitle.textContent = 'Edit Record';
        editFormContainer.classList.remove('active');
        return;
    }

    // Always make edit panel active and editable
    editFormContainer.classList.add('active');
    editActions.style.display = 'flex';
    editForm.innerHTML = '';

    // Use node.original for non-orphan records to access all row data
    const nodeData = node.isOrphan ? node.data : findNodeData(node.text.split(' - ')[0]);
    
    editPanelTitle.textContent = node.isOrphan ? 
        `Edit Orphan Record: ${node.text}` : 
        `Edit Record: ${node.text}`;
    
    if (nodeData) {
        editHistory = [{ ...nodeData }];
        
        const allColumns = Object.keys(excelData[0]).filter(field => 
            field !== 'Family Path' && field !== 'Parent-Child(s)'
        ); // Exclude summary columns

        allColumns.forEach(field => {
            const formGroup = document.createElement('div');
            formGroup.className = 'form-group';
            
            const label = document.createElement('label');
            label.textContent = field;
            
            const input = document.createElement('input');
            input.type = 'text';
            input.value = nodeData[field] || '';
            input.dataset.field = field;
            input.className = 'edit-input active';
            
            // Ensure input is always editable
            input.readOnly = false;
            input.disabled = false;
            
            formGroup.appendChild(label);
            formGroup.appendChild(input);
            editForm.appendChild(formGroup);
        });
    }
}

function findNodeData(nodeText) {
    // Ensure that the correct row is fetched based on child or parent mapping
    return excelData.find(row => 
        row[columnMappings.child] === nodeText || 
        row[columnMappings.parent] === nodeText
    );
}

function applyChanges() {
    if (!selectedNode || !excelData) return;

    const formInputs = document.querySelectorAll('#editForm input');
    const nodeText = selectedNode.text.split(' - ')[0];
    
    const rowIndex = excelData.findIndex(row => 
        row[columnMappings.parent] === nodeText || 
        row[columnMappings.child] === nodeText
    );

    if (rowIndex !== -1) {
        // Save current state before making changes
        editHistory.push({ ...excelData[rowIndex] });
        
        formInputs.forEach(input => {
            excelData[rowIndex][input.dataset.field] = input.value;
        });
        
        updateTreeView();
    }
}

function cancelChanges() {
    updateEditForm(selectedNode);
}

function updateTreeView() {
    if (!excelData || !columnMappings.parent || !columnMappings.child) return;

    // Update summaries before building tree
    updateSummaries();

    const treeData = [];
    const processedNodes = new Set();
    const parentChildMap = new Map();

    excelData.forEach(row => {
        const parentValue = row[columnMappings.parent];
        const childValue = row[columnMappings.child];
        const description = columnMappings.description ? row[columnMappings.description] : '';

        if (parentValue) {
            if (!parentChildMap.has(parentValue)) {
                parentChildMap.set(parentValue, []);
            }
            parentChildMap.get(parentValue).push({ child: childValue, description });
        }
    });

    function addNode(value, isParent = true, level = 0) {
        if (processedNodes.has(value)) return null;
        if (!value) return null;

        processedNodes.add(value);

        const icons = ['fas fa-folder', 'fas fa-folder-open', 'fas fa-toolbox'];
        const rowData = excelData.find(row => row[columnMappings.child] === value);
        const familyPath = rowData ? rowData['Family Path'] : 'N/A';
        const parentChild = rowData ? rowData['Parent-Child(s)'] : 'N/A';

        const node = {
            text: value,
            id: value,
            children: [],
            icon: isParent ? icons[0] : icons[2],
            data: rowData ? { ...rowData } : {} // Include all row data
        };

        if (parentChildMap.has(value)) {
            parentChildMap.get(value).forEach(({ child, description }) => {
                const childNode = addNode(child, false, level + 1);
                if (childNode) {
                    childNode.text = child + (description ? ` - ${description}` : '');
                    node.children.push(childNode);
                    node.icon = icons[1]; // Change to open folder if it has children
                }
            });
        }

        return node;
    }

    parentChildMap.forEach((_, parent) => {
        if (!processedNodes.has(parent)) {
            const node = addNode(parent);
            if (node) treeData.push(node);
        }
    });

    lastTreeData = treeData;
    $('#treeView').jstree(true).settings.core.data = treeData;
    $('#treeView').jstree(true).refresh();
}

function showOrphanRecords() {
    if (!excelData || !columnMappings.parent || !columnMappings.child) {
        alert('Please load data and map the Parent and Child columns first');
        return;
    }

    const orphans = excelData.filter(row => {
        return !row[columnMappings.parent] || 
               row[columnMappings.parent] === '' || 
               row[columnMappings.parent] === null;
    });

    if (orphans.length === 0) {
        alert('No orphan records found');
        return;
    }

    // Store orphans for searching
    window.currentOrphans = orphans;
    displayOrphanRecordsInModal(orphans);
}

function displayOrphanRecordsInModal(orphans) {
    const modal = document.getElementById('orphanRecordsModal');
    const orphansList = document.getElementById('orphansList');
    const editFormContainer = document.querySelector('.edit-form-container');
    
    orphansList.innerHTML = '';

    orphans.forEach(row => {
        const item = document.createElement('div');
        item.className = 'orphan-item';
        item.innerHTML = `
            <i class="fas fa-exclamation-circle"></i>
            <span>${row[columnMappings.child]}${columnMappings.description ? 
                ` - ${row[columnMappings.description] || ''}` : ''}</span>
        `;
        
        item.addEventListener('click', () => {
            // Set the selected node and update edit form
            selectedNode = { 
                text: row[columnMappings.child],
                data: row,
                isOrphan: true  // Flag to indicate this is an orphan record
            };
            
            // Ensure edit panel is active
            editFormContainer.classList.add('active');
            updateEditForm(selectedNode);
            
            // Highlight the selected orphan record
            orphansList.querySelectorAll('.orphan-item').forEach(i => i.classList.remove('selected'));
            item.classList.add('selected');
        });
        
        orphansList.appendChild(item);
    });

    modal.classList.add('show');
}

function initializeOrphanSearch() {
    const closeButton = document.getElementById('closeOrphanModal');
    const searchInput = document.getElementById('orphanModalSearch');

    closeButton.addEventListener('click', () => {
        document.getElementById('orphanRecordsModal').classList.remove('show');
    });

    searchInput.addEventListener('input', (e) => {
        const searchTerm = e.target.value.toLowerCase();
        if (!window.currentOrphans) return;

        const filteredOrphans = window.currentOrphans.filter(row => {
            const childMatch = row[columnMappings.child].toString().toLowerCase().includes(searchTerm);
            const descMatch = columnMappings.description && row[columnMappings.description] ?
                row[columnMappings.description].toString().toLowerCase().includes(searchTerm) : false;
            return childMatch || descMatch;
        });

        displayOrphanRecordsInModal(filteredOrphans);
    });
}

function saveChanges() {
    if (!excelData || !columnMappings.parent || !columnMappings.child) {
        alert('Please load data and map the Parent and Child columns first');
        return;
    }

    updateSummaries(); // Ensure summaries are up to date

    const ws = XLSX.utils.json_to_sheet(excelData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Updated Data");
    
    XLSX.writeFile(wb, 'updated_hierarchy.xlsx');
}

// Function to build Family Path for each row
function buildFamilyPaths() {
    const parentMap = new Map();
    excelData.forEach(row => {
        parentMap.set(row[columnMappings.child], row[columnMappings.parent]);
    });

    excelData.forEach(row => {
        let path = row[columnMappings.parent] ? row[columnMappings.parent] : row[columnMappings.child];
        let currentParent = parentMap.get(row[columnMappings.parent]);
        while (currentParent) {
            path = currentParent + ' > ' + path;
            currentParent = parentMap.get(currentParent);
        }
        row['Family Path'] = path;
    });
}

// Function to build Parent-Child(s) summary for each row
function buildParentChildSummary() {
    const childMap = new Map();
    const parentMap = new Map();

    excelData.forEach(row => {
        const parent = row[columnMappings.parent];
        const child = row[columnMappings.child];
        parentMap.set(child, parent);
        if (parent) {
            if (!childMap.has(parent)) {
                childMap.set(parent, []);
            }
            childMap.get(parent).push(child);
        }
    });

    excelData.forEach(row => {
        const parent = row[columnMappings.parent] || 'None';
        const children = childMap.get(row[columnMappings.child]) || [];
        row['Parent-Child(s)'] = `Parent: ${parent}; Children: ${children.join(', ')}`;
    });
}

// Function to update Family Path and Parent-Child(s) summaries
function updateSummaries() {
    buildFamilyPaths();
    buildParentChildSummary();
}

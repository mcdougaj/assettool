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
    initializeModals();
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

    // Store the original tree data if not already stored
    if (!window.originalTreeData) {
        window.originalTreeData = lastTreeData;
    }

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

    // Update the tree view with search results
    $('#treeView').jstree(true).settings.core.data = treeData;
    $('#treeView').jstree(true).refresh();
}

function clearSearch() {
    document.getElementById('searchInput').value = '';
    // Restore the original tree data
    if (window.originalTreeData) {
        $('#treeView').jstree(true).settings.core.data = window.originalTreeData;
        $('#treeView').jstree(true).refresh();
        window.originalTreeData = null;
    }
}

function buildTreeFromSearchResults(results) {
    // Build a tree structure from the search results
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

    function addNode(value) {
        if (processedNodes.has(value)) return null;
        processedNodes.add(value);

        const node = {
            text: value,
            id: value,
            children: [],
            icon: 'fas fa-folder'
        };

        if (parentChildMap.has(value)) {
            parentChildMap.get(value).forEach(({ child, description }) => {
                const childNode = addNode(child);
                if (childNode) {
                    childNode.text = child + (description ? ` - ${description}` : '');
                    node.children.push(childNode);
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
    
    updateTreeView();
}

function initializeTreeView() {
    $('#treeView').jstree({
        core: {
            check_callback: true, // Enable all modifications
            data: [],
            themes: {
                name: 'default',
                dots: true,
                icons: true,
                variant: 'large'
            }
        },
        plugins: ['dnd', 'wholerow'], // Include 'dnd' plugin for drag-and-drop
    })
    .on('select_node.jstree', function(e, data) {
        selectedNode = data.node;
        updateEditForm(selectedNode);
    })
    .on('move_node.jstree', function(e, data) {
        // Handle node movement to update parent in excelData

        // Get moved node's text (excluding description)
        const movedNodeText = data.node.text.split(' - ')[0];

        // Get new parent's text (excluding description)
        let newParentText;
        if (data.parent === '#') {
            // Moved to root level; parent is null or empty
            newParentText = '';
        } else {
            const parentNode = data.instance.get_node(data.parent);
            newParentText = parentNode.text.split(' - ')[0];
        }

        // Find the index of the moved node in excelData
        const rowIndex = excelData.findIndex(row =>
            row[columnMappings.child] === movedNodeText
        );

        if (rowIndex !== -1) {
            // Save current state before making changes
            historyStack.push(JSON.stringify(excelData));

            // Update the parent field in excelData
            excelData[rowIndex][columnMappings.parent] = newParentText || null;

            // Refresh the tree view to reflect changes
            updateTreeView();

            // Update the edit form with the moved node
            selectedNode = data.node;
            updateEditForm(selectedNode);
        }
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
    
    if (!node || !excelData) {
        editHistory = []; // Clear edit history for new node
        editForm.innerHTML = '<p class="placeholder-text">Select a node in the tree to edit its values</p>';
        editActions.style.display = 'none';
        return;
    }

    editForm.innerHTML = '';
    const nodeData = findNodeData(node.text.split(' - ')[0]);
    
    // Save initial state to edit history
    if (nodeData) {
        editHistory = [{ ...nodeData }];
    }

    const allColumns = Object.keys(excelData[0]);
    allColumns.forEach(field => {
        const formGroup = document.createElement('div');
        formGroup.className = 'form-group';
        
        const label = document.createElement('label');
        label.textContent = field;
        
        const input = document.createElement('input');
        input.type = 'text';
        input.value = nodeData ? nodeData[field] || '' : '';
        input.dataset.field = field;
        
        formGroup.appendChild(label);
        formGroup.appendChild(input);
        editForm.appendChild(formGroup);
    });
    
    editActions.style.display = 'flex';
}

function findNodeData(nodeText) {
    return excelData.find(row => {
        return row[columnMappings.parent] === nodeText || 
               row[columnMappings.child] === nodeText;
    });
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
        const node = {
            text: value,
            id: value,
            children: [],
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

    // Store the orphans data globally for searching
    window.currentOrphans = orphans;

    // Display all orphans initially
    searchOrphanRecords('');

    $('#orphanRecordsModal').addClass('show');
}

function displayOrphanRecords(orphans) {
    const treeData = orphans.map(row => ({
        text: row[columnMappings.child] + 
              (columnMappings.description && row[columnMappings.description] ? 
               ` - ${row[columnMappings.description]}` : ''),
        id: row[columnMappings.child],
        type: 'orphan',
        classes: 'orphan-record',
        icon: 'fas fa-exclamation-circle'
    }));

    $('#treeView').jstree(true).settings.core.data = treeData;
    $('#treeView').jstree(true).refresh();
}

function searchOrphans() {
    const searchTerm = document.getElementById('orphanSearchInput').value.toLowerCase();
    if (!window.currentOrphans) return;

    const filteredOrphans = window.currentOrphans.filter(row => {
        const childMatch = row[columnMappings.child].toString().toLowerCase().includes(searchTerm);
        const descMatch = columnMappings.description && row[columnMappings.description] ?
            row[columnMappings.description].toString().toLowerCase().includes(searchTerm) : false;
        return childMatch || descMatch;
    });

    displayOrphanRecords(filteredOrphans);
}

function clearOrphanSearch() {
    document.getElementById('orphanSearchInput').value = '';
    if (window.currentOrphans) {
        displayOrphanRecords(window.currentOrphans);
    }
}

// Add to initialization
function initializeOrphanSearch() {
    document.getElementById('orphanSearchButton').addEventListener('click', searchOrphans);
    document.getElementById('orphanSearchClear').addEventListener('click', clearOrphanSearch);
    document.getElementById('orphanSearchInput').addEventListener('keypress', function(e) {
        if (e.key === 'Enter') {
            searchOrphans();
        }
    });
}

function saveChanges() {
    if (!excelData || !columnMappings.parent || !columnMappings.child) {
        alert('Please load data and map the Parent and Child columns first');
        return;
    }

    const ws = XLSX.utils.json_to_sheet(excelData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Updated Data");
    
    XLSX.writeFile(wb, 'updated_hierarchy.xlsx');
}

function initializeModals() {
    // Initialize jsTree for orphan records
    $('#orphanRecordsTree').jstree({
        core: {
            themes: { name: 'default', dots: true, icons: true },
            data: []
        },
        plugins: ['search', 'wholerow']
    }).on('select_node.jstree', function(e, data) {
        // When an orphan record is selected, populate the Edit Records panel
        const nodeId = data.node.id;
        const orphanData = findOrphanData(nodeId);
        if (orphanData) {
            // Set the selectedNode to null to avoid conflicts with the main tree
            selectedNode = null;
            updateEditFormWithOrphan(orphanData);
            // Do not close the modal here
            // Scroll to the Edit Records panel if needed
            document.getElementById('rightPanel').scrollIntoView({ behavior: 'smooth' });
        }
    });

    // Initialize jsTree for search results
    $('#searchResultsTree').jstree({
        core: {
            themes: { name: 'default', dots: true, icons: true },
            data: []
        },
        plugins: ['search', 'wholerow']
    });

    // Orphan Records Modal events
    document.getElementById('closeOrphanModal').addEventListener('click', function() {
        $('#orphanRecordsModal').removeClass('show');
    });

    // Remove the search button event listener since we'll search on input
    // document.getElementById('orphanModalSearchButton').addEventListener('click', function() {
    //     const searchTerm = $('#orphanModalSearchInput').val();
    //     $('#orphanRecordsTree').jstree('search', searchTerm);
    // });

    // Add input event listener to search as user types
    document.getElementById('orphanModalSearchInput').addEventListener('input', function() {
        const searchTerm = this.value.toLowerCase();
        searchOrphanRecords(searchTerm);
    });

    // Optionally, remove the clear button if it's no longer needed
    // document.getElementById('orphanModalClearButton').addEventListener('click', function() {
    //     $('#orphanModalSearchInput').val('');
    //     $('#orphanRecordsTree').jstree('clear_search');
    // });

    // Search Results Modal events
    document.getElementById('closeSearchModal').addEventListener('click', function() {
        $('#searchResultsModal').removeClass('show');
    });

    document.getElementById('searchModalButton').addEventListener('click', function() {
        const searchTerm = $('#searchModalInput').val();
        $('#searchResultsTree').jstree('search', searchTerm);
    });

    document.getElementById('searchModalClear').addEventListener('click', function() {
        $('#searchModalInput').val('');
        $('#searchResultsTree').jstree('clear_search');
    });
}

function searchOrphanRecords(searchTerm) {
    if (!window.currentOrphans) return;

    const filteredOrphans = window.currentOrphans.filter(row => {
        const childValue = row[columnMappings.child]?.toString().toLowerCase() || '';
        const descriptionValue = columnMappings.description && row[columnMappings.description]
            ? row[columnMappings.description].toString().toLowerCase()
            : '';
        return childValue.includes(searchTerm) || descriptionValue.includes(searchTerm);
    });

    const treeData = filteredOrphans.map(row => ({
        text: row[columnMappings.child] +
            (columnMappings.description && row[columnMappings.description]
                ? ` - ${row[columnMappings.description]}`
                : ''),
        id: row[columnMappings.child],
        icon: 'fas fa-exclamation-circle'
    }));

    $('#orphanRecordsTree').jstree(true).settings.core.data = treeData;
    $('#orphanRecordsTree').jstree(true).refresh();
}

function findOrphanData(childId) {
    if (!window.currentOrphans) return null;
    return window.currentOrphans.find(row => row[columnMappings.child] === childId);
}

// Function to update the Edit Records panel with orphan data
function updateEditFormWithOrphan(orphanData) {
    const editForm = document.getElementById('editForm');
    const editActions = document.querySelector('.edit-actions');

    editForm.innerHTML = '';
    editHistory = [{ ...orphanData }]; // Save initial state

    const allColumns = Object.keys(excelData[0]);
    allColumns.forEach(field => {
        const formGroup = document.createElement('div');
        formGroup.className = 'form-group';
        
        const label = document.createElement('label');
        label.textContent = field;
        
        const input = document.createElement('input');
        input.type = 'text';
        input.value = orphanData[field] || '';
        input.dataset.field = field;
        
        formGroup.appendChild(label);
        formGroup.appendChild(input);
        editForm.appendChild(formGroup);
    });
    
    editActions.style.display = 'flex';
}

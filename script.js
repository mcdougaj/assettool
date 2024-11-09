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
let savedTreeState = null; // Add this near other global variables
let savedJsTreeState = null;
let isInSearchOrOrphanView = false;

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

// Replace the initializeSearch function
function initializeSearch() {
    const searchButton = document.getElementById('searchButton');
    const clearButton = document.getElementById('clearSearch');
    const searchInput = document.getElementById('searchInput');

    searchButton.addEventListener('click', performSearch);
    clearButton.addEventListener('click', handleBackButton);
    searchInput.addEventListener('keypress', function(e) {
        if (e.key === 'Enter') {
            performSearch();
        }
    });
}

// Modify the handleBackButton function
function handleBackButton() {
    if (isInSearchOrOrphanView && savedTreeState) {
        // Restore previous tree data and state
        $('#treeView').jstree(true).settings.core.data = savedTreeState;
        $('#treeView').jstree(true).refresh();
        $('#treeView').on('refresh.jstree', function() {
            $('#treeView').jstree(true).set_state(savedJsTreeState);
            $('#treeView').off('refresh.jstree');
        });
        
        // Reset states
        document.getElementById('orphanSearchContainer').style.display = 'none';
        document.getElementById('clearSearch').textContent = 'Reset View';
        window.currentOrphans = null;
        isInSearchOrOrphanView = false;
        savedTreeState = null;
        savedJsTreeState = null;
    } else {
        // Normal reset behavior
        updateTreeView();
    }
    document.getElementById('searchInput').value = '';
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

    if (searchResults.length > 0) {
        if (!isInSearchOrOrphanView) {
            savedTreeState = $('#treeView').jstree(true).settings.core.data;
            savedJsTreeState = $('#treeView').jstree(true).get_state();
            isInSearchOrOrphanView = true;
            document.getElementById('clearSearch').textContent = 'Back';
        }
        const treeData = buildTreeFromSearchResults(searchResults);
        $('#searchResultsTree').jstree(true).settings.core.data = treeData;
        $('#searchResultsTree').jstree(true).refresh();
        const modal = document.getElementById('searchResultsModal');
        const modalContent = modal.querySelector('.modal-content');
        resetModalPosition(modalContent);
        modal.classList.add('show');
    }
}

function buildTreeFromSearchResults(results) {
    const treeData = [];
    const processedNodes = new Set();
    const parentChildMap = new Map();

    results.forEach((row, index) => {
        const parentValue = row[columnMappings.parent];
        const childValue = row[columnMappings.child];
        const description = columnMappings.description ? row[columnMappings.description] : '';

        if (parentValue) {
            if (!parentChildMap.has(parentValue)) {
                parentChildMap.set(parentValue, []);
            }
            // Include row index
            parentChildMap.get(parentValue).push({ child: childValue, description, rowIndex: index });
        }
    });

    function addNode(value, isParent = true, level = 0, rowIndex = null) {
        if (processedNodes.has(value)) return null;
        if (!value) return null;
        
        processedNodes.add(value);

        const icons = ['fas fa-folder', 'fas fa-folder-open', 'fas fa-toolbox'];
        const node = {
            text: value,
            id: `${value}_${level}`,
            children: [],
            data: { rowIndex: rowIndex }, // Store row index
            state: { opened: true },
            icon: isParent ? icons[0] : icons[2]
        };

        if (parentChildMap.has(value)) {
            parentChildMap.get(value).forEach(({ child, description, rowIndex }) => {
                const childNode = addNode(child, false, level + 1, rowIndex);
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
    
    if (document.getElementById('orphanSearchContainer').style.display === 'block') {
        // If we're in orphan view, restore the normal hierarchy
        document.getElementById('orphanSearchContainer').style.display = 'none';
        window.currentOrphans = null;
        if (savedTreeState) {
            $('#treeView').jstree(true).settings.core.data = savedTreeState;
            $('#treeView').jstree(true).refresh();
        }
    } else {
        // Normal search clear behavior
        updateTreeView();
    }
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
    const rowIndex = node.data.rowIndex;

    if (rowIndex === undefined || rowIndex === null) {
        editForm.innerHTML = '<p class="placeholder-text">Unable to find data for the selected node.</p>';
        editActions.style.display = 'none';
        return;
    }

    const nodeData = excelData[rowIndex];

    // Save initial state to edit history
    if (nodeData) {
        editHistory = [{ ...nodeData }];
    } else {
        editForm.innerHTML = '<p class="placeholder-text">No data associated with this node.</p>';
        editActions.style.display = 'none';
        return;
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

    excelData.forEach((row, index) => {
        const parentValue = row[columnMappings.parent];
        const childValue = row[columnMappings.child];
        const description = columnMappings.description ? row[columnMappings.description] : '';

        if (parentValue) {
            if (!parentChildMap.has(parentValue)) {
                parentChildMap.set(parentValue, []);
            }
            // Include row index
            parentChildMap.get(parentValue).push({ child: childValue, description, rowIndex: index });
        }
    });

    function addNode(value, isParent = true, level = 0, rowIndex = null) {
        if (processedNodes.has(value)) return null;
        if (!value) return null;
        
        processedNodes.add(value);

        const icons = ['fas fa-folder', 'fas fa-folder-open', 'fas fa-toolbox'];
        const node = {
            text: value,
            id: `${value}_${level}`,
            children: [],
            data: { rowIndex: rowIndex }, // Store row index
            icon: isParent ? icons[0] : icons[2]
        };

        if (parentChildMap.has(value)) {
            parentChildMap.get(value).forEach(({ child, description, rowIndex }) => {
                const childNode = addNode(child, false, level + 1, rowIndex);
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

// Modify the showOrphanRecords function
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

    // Display orphans in modal
    const treeData = orphans.map((row, index) => ({
        text: row[columnMappings.child] + 
              (columnMappings.description && row[columnMappings.description] ? 
               ` - ${row[columnMappings.description]}` : ''),
        id: row[columnMappings.child],
        data: { rowIndex: index },
        icon: 'fas fa-exclamation-circle'
    }));

    $('#orphanRecordsTree').jstree(true).settings.core.data = treeData;
    $('#orphanRecordsTree').jstree(true).refresh();
    const modal = document.getElementById('orphanRecordsModal');
    const modalContent = modal.querySelector('.modal-content');
    resetModalPosition(modalContent);
    modal.classList.add('show');
    document.getElementById('orphanSearchContainer').style.display = 'block';
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

    const treeData = filteredOrphans.map((row, index) => ({
        text: row[columnMappings.child] + 
              (columnMappings.description && row[columnMappings.description] ? 
               ` - ${row[columnMappings.description]}` : ''),
        id: row[columnMappings.child],
        data: { rowIndex: window.currentOrphans.indexOf(row) },
        icon: 'fas fa-exclamation-circle'
    }));

    $('#orphanRecordsTree').jstree(true).settings.core.data = treeData;
    $('#orphanRecordsTree').jstree(true).refresh();
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

// Add these functions near the top with other initializations
// Modify the initializeModals function to setup tree node selection
function initializeModals() {
    document.getElementById('closeSearchResults').addEventListener('click', () => {
        document.getElementById('searchResultsModal').classList.remove('show');
    });
    
    document.getElementById('closeOrphanRecords').addEventListener('click', () => {
        document.getElementById('orphanRecordsModal').classList.remove('show');
        document.getElementById('orphanSearchContainer').style.display = 'none';
    });

    // Initialize the trees inside modals with node selection handling
    $('#searchResultsTree, #orphanRecordsTree').jstree({
        core: {
            themes: { name: 'default', dots: true, icons: true },
            data: []
        },
        plugins: ['wholerow', 'search'],
        search: {
            show_only_matches: true,
            show_only_matches_children: true,
            close_opened_onclear: false
        }
    }).on('select_node.jstree', function(e, data) {
        const rowIndex = findDataRowIndex(data.node);
        if (rowIndex !== -1) {
            selectedNode = data.node;
            const rightPanel = document.getElementById('rightPanel');
            rightPanel.style.display = 'flex';
            rightPanel.style.zIndex = '1001'; // Ensure it's above the modal overlay
            updateEditForm({ data: { rowIndex: rowIndex } });
        }
    });

    // Add modal search functionality
    const modalSearchInput = document.getElementById('modalSearchInput');
    const modalSearchClear = document.getElementById('modalSearchClear');
    const modalOrphanSearchInput = document.getElementById('modalOrphanSearchInput');
    const modalOrphanSearchClear = document.getElementById('modalOrphanSearchClear');

    modalSearchInput.addEventListener('input', function() {
        $('#searchResultsTree').jstree(true).search(this.value);
    });

    modalSearchClear.addEventListener('click', function() {
        modalSearchInput.value = '';
        $('#searchResultsTree').jstree(true).clear_search();
    });

    modalOrphanSearchInput.addEventListener('input', function() {
        $('#orphanRecordsTree').jstree(true).search(this.value);
    });

    modalOrphanSearchClear.addEventListener('click', function() {
        modalOrphanSearchInput.value = '';
        $('#orphanRecordsTree').jstree(true).clear_search();
    });

    initializeDraggableModals();
}

// Add helper function to find data row index
function findDataRowIndex(node) {
    const nodeText = node.text.split(' - ')[0]; // Get the text before any description
    return excelData.findIndex(row => 
        row[columnMappings.parent] === nodeText || 
        row[columnMappings.child] === nodeText
    );
}

// Replace the initializeDraggableModals function
function initializeDraggableModals() {
    const modals = document.querySelectorAll('.modal-content');
    
    modals.forEach(modal => {
        const header = modal.querySelector('.modal-header');
        let isDragging = false;
        let currentX = 0;
        let currentY = 0;
        let initialX = 0;
        let initialY = 0;

        header.addEventListener('mousedown', dragStart);
        document.addEventListener('mousemove', drag);
        document.addEventListener('mouseup', dragEnd);

        function dragStart(e) {
            if (e.target === header) {
                const rect = modal.getBoundingClientRect();
                initialX = e.clientX - rect.left;
                initialY = e.clientY - rect.top;
                
                isDragging = true;
                modal.classList.add('dragging');
            }
        }

        function drag(e) {
            if (!isDragging) return;

            e.preventDefault();
            currentX = e.clientX - initialX;
            currentY = e.clientY - initialY;

            // Keep modal within viewport bounds
            const rect = modal.getBoundingClientRect();
            const viewportWidth = window.innerWidth;
            const viewportHeight = window.innerHeight;

            if (currentX < 0) currentX = 0;
            if (currentY < 0) currentY = 0;
            if (currentX + rect.width > viewportWidth) currentX = viewportWidth - rect.width;
            if (currentY + rect.height > viewportHeight) currentY = viewportHeight - rect.height;

            modal.style.left = `${currentX}px`;
            modal.style.top = `${currentY}px`;
            modal.style.transform = 'none';
        }

        function dragEnd() {
            isDragging = false;
            modal.classList.remove('dragging');
        }
    });
}

// Update the resetModalPosition function
function resetModalPosition(modalContent) {
    modalContent.style.left = '50%';
    modalContent.style.top = '50%';
    modalContent.style.transform = 'translate(-50%, -50%)';
}
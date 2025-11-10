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
let selectedNodes = []; // For multi-select
let pmRecords = []; // PM records storage
let bomRecords = []; // BOM records storage
let assetPmAssignments = {}; // Maps asset IDs to PM records
let assetBomAssignments = {}; // Maps asset IDs to BOM records

// Constants
const MAX_FILE_SIZE = 10 * 1024 * 1024; // 10MB
const TOAST_DURATION = 4000; // 4 seconds

// Utility Functions
function sanitizeInput(input) {
    if (typeof input !== 'string') return input;
    const div = document.createElement('div');
    div.textContent = input;
    return div.innerHTML;
}

function showToast(message, type = 'info') {
    const container = document.getElementById('toastContainer');
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;

    const icons = {
        success: '✓',
        error: '✕',
        warning: '⚠',
        info: 'ℹ'
    };

    toast.innerHTML = `
        <span class="toast-icon">${icons[type] || icons.info}</span>
        <span class="toast-message">${sanitizeInput(message)}</span>
        <button class="toast-close">×</button>
    `;

    container.appendChild(toast);

    const closeBtn = toast.querySelector('.toast-close');
    closeBtn.addEventListener('click', () => removeToast(toast));

    setTimeout(() => removeToast(toast), TOAST_DURATION);
}

function removeToast(toast) {
    toast.classList.add('hiding');
    setTimeout(() => {
        if (toast.parentNode) {
            toast.parentNode.removeChild(toast);
        }
    }, 300);
}

function showConfirmDialog(title, message) {
    return new Promise((resolve) => {
        const dialog = document.getElementById('confirmDialog');
        const titleEl = document.getElementById('confirmTitle');
        const messageEl = document.getElementById('confirmMessage');
        const yesBtn = document.getElementById('confirmYes');
        const noBtn = document.getElementById('confirmNo');

        titleEl.textContent = title;
        messageEl.textContent = message;
        dialog.classList.add('show');

        function handleYes() {
            dialog.classList.remove('show');
            cleanup();
            resolve(true);
        }

        function handleNo() {
            dialog.classList.remove('show');
            cleanup();
            resolve(false);
        }

        function cleanup() {
            yesBtn.removeEventListener('click', handleYes);
            noBtn.removeEventListener('click', handleNo);
        }

        yesBtn.addEventListener('click', handleYes);
        noBtn.addEventListener('click', handleNo);
    });
}

function showLoading(message = 'Processing...') {
    const overlay = document.createElement('div');
    overlay.id = 'loadingOverlay';
    overlay.className = 'loading-overlay';
    overlay.innerHTML = `
        <div class="loading-spinner"></div>
        <div class="loading-text">${sanitizeInput(message)}</div>
    `;
    document.body.appendChild(overlay);
}

function hideLoading() {
    const overlay = document.getElementById('loadingOverlay');
    if (overlay) {
        overlay.remove();
    }
}

function validateFileSize(file) {
    if (file.size > MAX_FILE_SIZE) {
        showToast(`File size exceeds ${MAX_FILE_SIZE / (1024 * 1024)}MB limit`, 'error');
        return false;
    }
    return true;
}

function getSelectedTreeNodes() {
    const tree = $('#treeView').jstree(true);
    const selected = tree.get_selected(true);
    return selected;
}

// Get all child nodes recursively for a given node
function getAllChildNodes(nodeId) {
    const tree = $('#treeView').jstree(true);
    const node = tree.get_node(nodeId);
    const children = [];

    function collectChildren(n) {
        if (n.children && n.children.length > 0) {
            n.children.forEach(childId => {
                const childNode = tree.get_node(childId);
                children.push(childNode);
                collectChildren(childNode);
            });
        }
    }

    collectChildren(node);
    return children;
}

// Get all nodes that will be affected (includes children if parent selected)
function getAffectedNodes(selectedNodes) {
    const affected = new Set();

    selectedNodes.forEach(node => {
        affected.add(node.id);
        const children = getAllChildNodes(node.id);
        children.forEach(child => affected.add(child.id));
    });

    return Array.from(affected);
}

// Get all assets (leaf nodes and parents) for the copy dropdown
function getAllAssetNodes() {
    const tree = $('#treeView').jstree(true);
    if (!tree) return [];

    const allNodes = tree.get_json('#', { flat: true });
    return allNodes.filter(node => node.id !== '#');
}

// Initialize sample PM and BOM data
function initializeSampleData() {
    pmRecords = [
        { id: 'PM001', code: 'PM-DAILY-001', description: 'Daily Visual Inspection', frequency: 'Daily', type: 'Inspection' },
        { id: 'PM002', code: 'PM-WEEKLY-001', description: 'Weekly Lubrication Check', frequency: 'Weekly', type: 'Maintenance' },
        { id: 'PM003', code: 'PM-MONTHLY-001', description: 'Monthly Calibration', frequency: 'Monthly', type: 'Calibration' },
        { id: 'PM004', code: 'PM-QUARTERLY-001', description: 'Quarterly Safety Inspection', frequency: 'Quarterly', type: 'Safety' },
        { id: 'PM005', code: 'PM-ANNUAL-001', description: 'Annual Equipment Overhaul', frequency: 'Annual', type: 'Overhaul' }
    ];

    bomRecords = [
        { id: 'BOM001', partNumber: 'PART-001', description: 'Hydraulic Oil Filter', quantity: 2, unit: 'pcs' },
        { id: 'BOM002', partNumber: 'PART-002', description: 'Air Filter Element', quantity: 1, unit: 'pcs' },
        { id: 'BOM003', partNumber: 'PART-003', description: 'Bearing Assembly', quantity: 4, unit: 'pcs' },
        { id: 'BOM004', partNumber: 'PART-004', description: 'Drive Belt', quantity: 2, unit: 'pcs' },
        { id: 'BOM005', partNumber: 'PART-005', description: 'Lubrication Grease', quantity: 1, unit: 'kg' }
    ];
}

document.addEventListener('DOMContentLoaded', function() {
    initializeSampleData();
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
    initializePMManagement();
    initializeBOMManagement();
});

function initializeUndoButton() {
    document.getElementById('undoHierarchy').addEventListener('click', undoLastChange);
}

function undoLastChange() {
    if (historyStack.length > 0) {
        excelData = JSON.parse(historyStack.pop()); // Restore previous state
        updateTreeView();
        showToast('Changes undone successfully', 'success');
    } else {
        showToast('No more changes to undo', 'info');
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
        try {
            const fileInput = document.getElementById('excelFile');
            const file = fileInput.files[0];
            if (!file) {
                showToast('Please select a file first', 'warning');
                return;
            }

            if (!validateFileSize(file)) {
                return;
            }

            showLoading('Loading file...');
            const reader = new FileReader();

            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    originalWorkbook = XLSX.read(data, { type: 'array' });
                    const firstSheet = originalWorkbook.Sheets[originalWorkbook.SheetNames[0]];
                    excelData = XLSX.utils.sheet_to_json(firstSheet);
                    populateColumnsList();
                    hideLoading();
                    showToast('File loaded successfully', 'success');
                } catch (error) {
                    hideLoading();
                    console.error('Error reading file:', error);
                    showToast('Failed to read Excel file. Please check the file format', 'error');
                }
            };

            reader.onerror = function() {
                hideLoading();
                showToast('Failed to read file', 'error');
            };

            reader.readAsArrayBuffer(file);
        } catch (error) {
            hideLoading();
            console.error('Error loading file:', error);
            showToast('An error occurred while loading the file', 'error');
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
            showToast('Please map the Parent and Child columns first', 'warning');
            return;
        }
        modal.classList.add('show');
    });

    saveButton.addEventListener('click', function() {
        try {
            const parentValue = sanitizeInput(document.getElementById('newParentValue').value.trim());
            const childValue = sanitizeInput(document.getElementById('newChildValue').value.trim());
            const description = sanitizeInput(document.getElementById('newDescription').value.trim());

            if (!parentValue || !childValue) {
                showToast('Parent and Child values are required', 'warning');
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
            showToast('New record added successfully', 'success');
        } catch (error) {
            console.error('Error adding new record:', error);
            showToast('Failed to add new record', 'error');
        }
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
        showToast('No matches found', 'info');
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

        // Handle different node types
        if (data.node.type === 'pm-item' || data.node.type === 'bom-item') {
            // Show PM/BOM item details
            updatePMBOMItemDetails(selectedNode);
        } else if (data.node.type === 'pm-folder' || data.node.type === 'bom-folder') {
            // For folders, just show a summary - don't allow editing
            showFolderSummary(selectedNode);
        } else {
            // Regular asset node
            updateEditForm(selectedNode);
            updatePMBOMDetails(selectedNode);
        }
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
            showToast('Edit changes undone', 'success');
        }
    } else {
        showToast('No more edit changes to undo', 'info');
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

        // Get PM and BOM counts for this node
        const pmCount = assetPmAssignments[value] ? assetPmAssignments[value].length : 0;
        const bomCount = assetBomAssignments[value] ? assetBomAssignments[value].length : 0;

        // Build node text with badges
        let nodeText = value;
        if (pmCount > 0) {
            nodeText += ` <span class="tree-badge pm-badge">PM: ${pmCount}</span>`;
        }
        if (bomCount > 0) {
            nodeText += ` <span class="tree-badge bom-badge">BOM: ${bomCount}</span>`;
        }

        const node = {
            text: nodeText,
            id: value,
            children: [],
            icon: isParent ? icons[0] : icons[2],
            data: rowData ? { ...rowData } : {}, // Include all row data
            type: 'asset' // Mark as asset node
        };

        // Add regular child nodes first
        if (parentChildMap.has(value)) {
            parentChildMap.get(value).forEach(({ child, description }) => {
                const childNode = addNode(child, false, level + 1);
                if (childNode) {
                    const childPmCount = assetPmAssignments[child] ? assetPmAssignments[child].length : 0;
                    const childBomCount = assetBomAssignments[child] ? assetBomAssignments[child].length : 0;

                    let childText = child + (description ? ` - ${description}` : '');
                    if (childPmCount > 0) {
                        childText += ` <span class="tree-badge pm-badge">PM: ${childPmCount}</span>`;
                    }
                    if (childBomCount > 0) {
                        childText += ` <span class="tree-badge bom-badge">BOM: ${childBomCount}</span>`;
                    }

                    childNode.text = childText;
                    node.children.push(childNode);
                    node.icon = icons[1]; // Change to open folder if it has children
                }
            });
        }

        // Add PM Tasks folder if there are PM assignments
        if (pmCount > 0) {
            const pmFolderNode = {
                text: `<i class="fas fa-clipboard-check"></i> PM Tasks (${pmCount})`,
                id: `${value}_pm_folder`,
                children: [],
                icon: 'fas fa-clipboard-list',
                type: 'pm-folder',
                state: { opened: false }
            };

            const pmAssignments = assetPmAssignments[value] || [];
            pmAssignments.forEach(pmId => {
                const pm = pmRecords.find(p => p.id === pmId);
                if (pm) {
                    pmFolderNode.children.push({
                        text: `${sanitizeInput(pm.code)} - ${sanitizeInput(pm.description)}`,
                        id: `${value}_pm_${pmId}`,
                        icon: 'fas fa-tasks',
                        type: 'pm-item',
                        data: { pmId: pmId, assetId: value, pm: pm }
                    });
                }
            });

            node.children.push(pmFolderNode);
            node.icon = icons[1]; // Change to open folder
        }

        // Add BOM Items folder if there are BOM assignments
        if (bomCount > 0) {
            const bomFolderNode = {
                text: `<i class="fas fa-box"></i> BOM Items (${bomCount})`,
                id: `${value}_bom_folder`,
                children: [],
                icon: 'fas fa-boxes',
                type: 'bom-folder',
                state: { opened: false }
            };

            const bomAssignments = assetBomAssignments[value] || [];
            bomAssignments.forEach(bomId => {
                const bom = bomRecords.find(b => b.id === bomId);
                if (bom) {
                    bomFolderNode.children.push({
                        text: `${sanitizeInput(bom.partNumber)} - ${sanitizeInput(bom.description)} (Qty: ${bom.quantity})`,
                        id: `${value}_bom_${bomId}`,
                        icon: 'fas fa-cube',
                        type: 'bom-item',
                        data: { bomId: bomId, assetId: value, bom: bom }
                    });
                }
            });

            node.children.push(bomFolderNode);
            node.icon = icons[1]; // Change to open folder
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
    try {
        if (!excelData || !columnMappings.parent || !columnMappings.child) {
            showToast('Please load data and map the Parent and Child columns first', 'warning');
            return;
        }

        const orphans = excelData.filter(row => {
            return !row[columnMappings.parent] ||
                   row[columnMappings.parent] === '' ||
                   row[columnMappings.parent] === null;
        });

        if (orphans.length === 0) {
            showToast('No orphan records found', 'info');
            return;
        }

        // Store orphans for searching
        window.currentOrphans = orphans;
        displayOrphanRecordsInModal(orphans);
    } catch (error) {
        console.error('Error showing orphan records:', error);
        showToast('Failed to load orphan records', 'error');
    }
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

async function saveChanges() {
    try {
        if (!excelData || !columnMappings.parent || !columnMappings.child) {
            showToast('Please load data and map the Parent and Child columns first', 'warning');
            return;
        }

        const confirmed = await showConfirmDialog(
            'Export Changes',
            'Are you sure you want to export the current hierarchy to Excel?'
        );

        if (!confirmed) {
            return;
        }

        showLoading('Exporting data...');
        updateSummaries(); // Ensure summaries are up to date

        // Add PM and BOM data to export
        const exportData = excelData.map(row => {
            const nodeId = row[columnMappings.child];
            const pmList = assetPmAssignments[nodeId] || [];
            const bomList = assetBomAssignments[nodeId] || [];

            return {
                ...row,
                'PM Tasks': pmList.map(pmId => {
                    const pm = pmRecords.find(p => p.id === pmId);
                    return pm ? pm.code : '';
                }).join(', '),
                'BOM Items': bomList.map(bomId => {
                    const bom = bomRecords.find(b => b.id === bomId);
                    return bom ? bom.partNumber : '';
                }).join(', ')
            };
        });

        const ws = XLSX.utils.json_to_sheet(exportData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Updated Data");

        XLSX.writeFile(wb, 'updated_hierarchy.xlsx');
        hideLoading();
        showToast('Data exported successfully', 'success');
    } catch (error) {
        hideLoading();
        console.error('Error saving changes:', error);
        showToast('Failed to export data', 'error');
    }
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

// PM Management Functions
function initializePMManagement() {
    const managePmBtn = document.getElementById('managePM');
    const closePmBtn = document.getElementById('closePmModal');
    const closePmBtnX = document.getElementById('closePmModalX');
    const addPmBtn = document.getElementById('addNewPm');
    const loadPmFileBtn = document.getElementById('loadPmFile');
    const pmSearchInput = document.getElementById('pmSearchInput');
    const copyPmBtn = document.getElementById('copyPmBtn');

    managePmBtn.addEventListener('click', openPMManagementModal);
    closePmBtn.addEventListener('click', closePMManagementModal);
    closePmBtnX.addEventListener('click', closePMManagementModal);
    addPmBtn.addEventListener('click', addNewPMRecord);
    loadPmFileBtn.addEventListener('click', loadPMFromFile);
    pmSearchInput.addEventListener('input', filterPMList);
    copyPmBtn.addEventListener('click', copyPMFromAsset);
}

function openPMManagementModal() {
    try {
        selectedNodes = getSelectedTreeNodes();
        if (selectedNodes.length === 0) {
            showToast('Please select at least one node in the tree', 'warning');
            return;
        }

        const modal = document.getElementById('pmManagementModal');
        modal.classList.add('show');

        displaySelectedNodes('selectedNodesDisplay');
        populateCopySourceDropdown('pmCopySource');
        updateAffectedNodesCount('affectedNodesCount', 'affectedNodesInfo');
        renderPMList();
        renderAssignedPMs();
    } catch (error) {
        console.error('Error opening PM management:', error);
        showToast('Failed to open PM management', 'error');
    }
}

function closePMManagementModal() {
    const modal = document.getElementById('pmManagementModal');
    modal.classList.remove('show');
    selectedNodes = [];
}

function renderPMList(filter = '') {
    const pmList = document.getElementById('pmList');
    pmList.innerHTML = '';

    const filteredPMs = pmRecords.filter(pm =>
        pm.code.toLowerCase().includes(filter.toLowerCase()) ||
        pm.description.toLowerCase().includes(filter.toLowerCase())
    );

    if (filteredPMs.length === 0) {
        pmList.innerHTML = '<p class="placeholder-text">No PM records found</p>';
        return;
    }

    filteredPMs.forEach(pm => {
        const card = document.createElement('div');
        card.className = 'item-card';

        // Check if this PM is already assigned to any selected node
        const isAssigned = selectedNodes.some(node =>
            assetPmAssignments[node.id] && assetPmAssignments[node.id].includes(pm.id)
        );

        card.innerHTML = `
            <div class="item-info">
                <div class="item-code">${sanitizeInput(pm.code)}</div>
                <div class="item-description">${sanitizeInput(pm.description)}</div>
                <div class="item-meta">Frequency: ${sanitizeInput(pm.frequency)} | Type: ${sanitizeInput(pm.type)}</div>
            </div>
            <button class="item-add ${isAssigned ? 'disabled' : ''}" data-pm-id="${pm.id}" ${isAssigned ? 'disabled' : ''}>
                ${isAssigned ? 'Added' : 'Add'}
            </button>
        `;

        if (!isAssigned) {
            const addBtn = card.querySelector('.item-add');
            addBtn.addEventListener('click', () => addPMToSelectedNodes(pm.id));
        }

        pmList.appendChild(card);
    });
}

function renderAssignedPMs() {
    const assignedList = document.getElementById('assignedPmList');
    assignedList.innerHTML = '';

    if (selectedNodes.length === 0) return;

    const nodeId = selectedNodes[0].id;
    const assigned = assetPmAssignments[nodeId] || [];

    if (assigned.length === 0) {
        assignedList.innerHTML = '<p class="placeholder-text">No PM tasks assigned</p>';
        return;
    }

    assigned.forEach(pmId => {
        const pm = pmRecords.find(p => p.id === pmId);
        if (!pm) return;

        const card = document.createElement('div');
        card.className = 'item-card';
        card.innerHTML = `
            <div class="item-info">
                <div class="item-code">${sanitizeInput(pm.code)}</div>
                <div class="item-description">${sanitizeInput(pm.description)}</div>
                <div class="item-meta">Frequency: ${sanitizeInput(pm.frequency)}</div>
            </div>
            <button class="item-remove" data-pm-id="${pm.id}">Remove</button>
        `;

        const removeBtn = card.querySelector('.item-remove');
        removeBtn.addEventListener('click', () => removePMAssignment(nodeId, pmId));

        assignedList.appendChild(card);
    });
}

function addPMToSelectedNodes(pmId) {
    try {
        // Get all affected nodes (includes children of parent nodes)
        const affectedNodeIds = getAffectedNodes(selectedNodes);

        let assignedCount = 0;
        affectedNodeIds.forEach(nodeId => {
            if (!assetPmAssignments[nodeId]) {
                assetPmAssignments[nodeId] = [];
            }

            if (!assetPmAssignments[nodeId].includes(pmId)) {
                assetPmAssignments[nodeId].push(pmId);
                assignedCount++;
            }
        });

        renderAssignedPMs();
        renderPMList(document.getElementById('pmSearchInput').value); // Refresh the list to update button states
        updateTreeView(); // Refresh tree to show updated badges

        const pm = pmRecords.find(p => p.id === pmId);
        if (pm) {
            showToast(`"${pm.code}" added to ${affectedNodeIds.length} asset(s)`, 'success');
        }
    } catch (error) {
        console.error('Error adding PM assignment:', error);
        showToast('Failed to add PM assignment', 'error');
    }
}

function removePMAssignment(nodeId, pmId) {
    if (assetPmAssignments[nodeId]) {
        assetPmAssignments[nodeId] = assetPmAssignments[nodeId].filter(id => id !== pmId);
        renderAssignedPMs();
        renderPMList(document.getElementById('pmSearchInput').value); // Refresh list to update button states
        updateTreeView(); // Refresh tree to show updated badges
        showToast('PM assignment removed', 'success');
    }
}

function addNewPMRecord() {
    try {
        const code = document.getElementById('newPmCode').value.trim();
        const description = document.getElementById('newPmDescription').value.trim();
        const frequency = document.getElementById('newPmFrequency').value.trim();

        if (!code || !description || !frequency) {
            showToast('Please fill in all PM fields', 'warning');
            return;
        }

        const newPM = {
            id: 'PM' + (Date.now()),
            code: sanitizeInput(code),
            description: sanitizeInput(description),
            frequency: sanitizeInput(frequency),
            type: 'Custom'
        };

        pmRecords.push(newPM);
        renderPMList();

        document.getElementById('newPmCode').value = '';
        document.getElementById('newPmDescription').value = '';
        document.getElementById('newPmFrequency').value = '';

        showToast('New PM record added', 'success');
    } catch (error) {
        console.error('Error adding PM record:', error);
        showToast('Failed to add PM record', 'error');
    }
}

function filterPMList() {
    const filter = document.getElementById('pmSearchInput').value;
    renderPMList(filter);
}

function loadPMFromFile() {
    try {
        const fileInput = document.getElementById('pmFileInput');
        const file = fileInput.files[0];

        if (!file) {
            showToast('Please select a file first', 'warning');
            return;
        }

        if (!validateFileSize(file)) {
            return;
        }

        showLoading('Loading PM records...');
        const reader = new FileReader();

        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(firstSheet);

                // Expected columns: Code, Description, Frequency, Type
                let loadedCount = 0;
                jsonData.forEach(row => {
                    if (row.Code && row.Description && row.Frequency) {
                        const newPM = {
                            id: 'PM' + Date.now() + '_' + loadedCount,
                            code: sanitizeInput(row.Code.toString()),
                            description: sanitizeInput(row.Description.toString()),
                            frequency: sanitizeInput(row.Frequency.toString()),
                            type: row.Type ? sanitizeInput(row.Type.toString()) : 'Imported'
                        };
                        pmRecords.push(newPM);
                        loadedCount++;
                    }
                });

                renderPMList();
                fileInput.value = ''; // Clear file input
                hideLoading();
                showToast(`Loaded ${loadedCount} PM records successfully`, 'success');
            } catch (error) {
                hideLoading();
                console.error('Error reading PM file:', error);
                showToast('Failed to read PM file. Please check the format (Code, Description, Frequency, Type columns required)', 'error');
            }
        };

        reader.onerror = function() {
            hideLoading();
            showToast('Failed to read file', 'error');
        };

        reader.readAsArrayBuffer(file);
    } catch (error) {
        hideLoading();
        console.error('Error loading PM file:', error);
        showToast('An error occurred while loading the PM file', 'error');
    }
}

// Helper function to populate copy source dropdown
function populateCopySourceDropdown(dropdownId) {
    const dropdown = document.getElementById(dropdownId);
    dropdown.innerHTML = '<option value="">-- Select Asset --</option>';

    const allNodes = getAllAssetNodes();
    allNodes.forEach(node => {
        const option = document.createElement('option');
        option.value = node.id;
        option.textContent = node.text;
        dropdown.appendChild(option);
    });
}

// Helper function to update affected nodes count display
function updateAffectedNodesCount(countElementId, infoElementId) {
    const affectedNodeIds = getAffectedNodes(selectedNodes);
    const countElement = document.getElementById(countElementId);
    const infoElement = document.getElementById(infoElementId);

    if (countElement) {
        countElement.textContent = affectedNodeIds.length;
    }

    if (infoElement) {
        if (affectedNodeIds.length > selectedNodes.length) {
            infoElement.style.display = 'block';
        } else {
            infoElement.style.display = 'none';
        }
    }
}

// Copy PM assignments from source asset to selected assets
function copyPMFromAsset() {
    try {
        const sourceNodeId = document.getElementById('pmCopySource').value;
        if (!sourceNodeId) {
            showToast('Please select a source asset to copy from', 'warning');
            return;
        }

        const sourcePMs = assetPmAssignments[sourceNodeId] || [];
        if (sourcePMs.length === 0) {
            showToast('Selected asset has no PM tasks assigned', 'info');
            return;
        }

        const affectedNodeIds = getAffectedNodes(selectedNodes);
        affectedNodeIds.forEach(nodeId => {
            if (!assetPmAssignments[nodeId]) {
                assetPmAssignments[nodeId] = [];
            }
            // Copy all PMs from source, avoiding duplicates
            sourcePMs.forEach(pmId => {
                if (!assetPmAssignments[nodeId].includes(pmId)) {
                    assetPmAssignments[nodeId].push(pmId);
                }
            });
        });

        renderAssignedPMs();
        renderPMList(document.getElementById('pmSearchInput').value);
        updateTreeView(); // Refresh tree to show updated badges

        const tree = $('#treeView').jstree(true);
        const sourceNode = tree.get_node(sourceNodeId);
        showToast(`Copied ${sourcePMs.length} PM task(s) from "${sourceNode.text}" to ${affectedNodeIds.length} asset(s)`, 'success');
    } catch (error) {
        console.error('Error copying PM assignments:', error);
        showToast('Failed to copy PM assignments', 'error');
    }
}

// BOM Management Functions
function initializeBOMManagement() {
    const manageBomBtn = document.getElementById('manageBOM');
    const closeBomBtn = document.getElementById('closeBomModal');
    const closeBomBtnX = document.getElementById('closeBomModalX');
    const addBomBtn = document.getElementById('addNewBom');
    const loadBomFileBtn = document.getElementById('loadBomFile');
    const bomSearchInput = document.getElementById('bomSearchInput');
    const copyBomBtn = document.getElementById('copyBomBtn');

    manageBomBtn.addEventListener('click', openBOMManagementModal);
    closeBomBtn.addEventListener('click', closeBOMManagementModal);
    closeBomBtnX.addEventListener('click', closeBOMManagementModal);
    addBomBtn.addEventListener('click', addNewBOMRecord);
    loadBomFileBtn.addEventListener('click', loadBOMFromFile);
    bomSearchInput.addEventListener('input', filterBOMList);
    copyBomBtn.addEventListener('click', copyBOMFromAsset);
}

function openBOMManagementModal() {
    try {
        selectedNodes = getSelectedTreeNodes();
        if (selectedNodes.length === 0) {
            showToast('Please select at least one node in the tree', 'warning');
            return;
        }

        const modal = document.getElementById('bomManagementModal');
        modal.classList.add('show');

        displaySelectedNodes('selectedNodesBomDisplay');
        populateCopySourceDropdown('bomCopySource');
        updateAffectedNodesCount('affectedNodesBomCount', 'affectedNodesBomInfo');
        renderBOMList();
        renderAssignedBOMs();
    } catch (error) {
        console.error('Error opening BOM management:', error);
        showToast('Failed to open BOM management', 'error');
    }
}

function closeBOMManagementModal() {
    const modal = document.getElementById('bomManagementModal');
    modal.classList.remove('show');
    selectedNodes = [];
}

function renderBOMList(filter = '') {
    const bomList = document.getElementById('bomList');
    bomList.innerHTML = '';

    const filteredBOMs = bomRecords.filter(bom =>
        bom.partNumber.toLowerCase().includes(filter.toLowerCase()) ||
        bom.description.toLowerCase().includes(filter.toLowerCase())
    );

    if (filteredBOMs.length === 0) {
        bomList.innerHTML = '<p class="placeholder-text">No inventory items found</p>';
        return;
    }

    filteredBOMs.forEach(bom => {
        const card = document.createElement('div');
        card.className = 'item-card';

        // Check if this BOM is already assigned to any selected node
        const isAssigned = selectedNodes.some(node =>
            assetBomAssignments[node.id] && assetBomAssignments[node.id].includes(bom.id)
        );

        card.innerHTML = `
            <div class="item-info">
                <div class="item-code">${sanitizeInput(bom.partNumber)}</div>
                <div class="item-description">${sanitizeInput(bom.description)}</div>
                <div class="item-meta">Quantity: ${bom.quantity} ${sanitizeInput(bom.unit)}</div>
            </div>
            <button class="item-add ${isAssigned ? 'disabled' : ''}" data-bom-id="${bom.id}" ${isAssigned ? 'disabled' : ''}>
                ${isAssigned ? 'Added' : 'Add'}
            </button>
        `;

        if (!isAssigned) {
            const addBtn = card.querySelector('.item-add');
            addBtn.addEventListener('click', () => addBOMToSelectedNodes(bom.id));
        }

        bomList.appendChild(card);
    });
}

function renderAssignedBOMs() {
    const assignedList = document.getElementById('assignedBomList');
    assignedList.innerHTML = '';

    if (selectedNodes.length === 0) return;

    const nodeId = selectedNodes[0].id;
    const assigned = assetBomAssignments[nodeId] || [];

    if (assigned.length === 0) {
        assignedList.innerHTML = '<p class="placeholder-text">No BOM items assigned</p>';
        return;
    }

    assigned.forEach(bomId => {
        const bom = bomRecords.find(b => b.id === bomId);
        if (!bom) return;

        const card = document.createElement('div');
        card.className = 'item-card';
        card.innerHTML = `
            <div class="item-info">
                <div class="item-code">${sanitizeInput(bom.partNumber)}</div>
                <div class="item-description">${sanitizeInput(bom.description)}</div>
                <div class="item-meta">Quantity: ${bom.quantity} ${sanitizeInput(bom.unit)}</div>
            </div>
            <button class="item-remove" data-bom-id="${bom.id}">Remove</button>
        `;

        const removeBtn = card.querySelector('.item-remove');
        removeBtn.addEventListener('click', () => removeBOMAssignment(nodeId, bomId));

        assignedList.appendChild(card);
    });
}

function addBOMToSelectedNodes(bomId) {
    try {
        // Get all affected nodes (includes children of parent nodes)
        const affectedNodeIds = getAffectedNodes(selectedNodes);

        let assignedCount = 0;
        affectedNodeIds.forEach(nodeId => {
            if (!assetBomAssignments[nodeId]) {
                assetBomAssignments[nodeId] = [];
            }

            if (!assetBomAssignments[nodeId].includes(bomId)) {
                assetBomAssignments[nodeId].push(bomId);
                assignedCount++;
            }
        });

        renderAssignedBOMs();
        renderBOMList(document.getElementById('bomSearchInput').value); // Refresh the list to update button states
        updateTreeView(); // Refresh tree to show updated badges

        const bom = bomRecords.find(b => b.id === bomId);
        if (bom) {
            showToast(`"${bom.partNumber}" added to ${affectedNodeIds.length} asset(s)`, 'success');
        }
    } catch (error) {
        console.error('Error adding BOM assignment:', error);
        showToast('Failed to add BOM assignment', 'error');
    }
}

function removeBOMAssignment(nodeId, bomId) {
    if (assetBomAssignments[nodeId]) {
        assetBomAssignments[nodeId] = assetBomAssignments[nodeId].filter(id => id !== bomId);
        renderAssignedBOMs();
        renderBOMList(document.getElementById('bomSearchInput').value); // Refresh list to update button states
        updateTreeView(); // Refresh tree to show updated badges
        showToast('BOM assignment removed', 'success');
    }
}

function addNewBOMRecord() {
    try {
        const partNumber = document.getElementById('newBomPartNumber').value.trim();
        const description = document.getElementById('newBomDescription').value.trim();
        const quantity = document.getElementById('newBomQuantity').value;

        if (!partNumber || !description || !quantity) {
            showToast('Please fill in all BOM fields', 'warning');
            return;
        }

        const newBOM = {
            id: 'BOM' + (Date.now()),
            partNumber: sanitizeInput(partNumber),
            description: sanitizeInput(description),
            quantity: parseInt(quantity),
            unit: 'pcs'
        };

        bomRecords.push(newBOM);
        renderBOMList();

        document.getElementById('newBomPartNumber').value = '';
        document.getElementById('newBomDescription').value = '';
        document.getElementById('newBomQuantity').value = '1';

        showToast('New BOM item added', 'success');
    } catch (error) {
        console.error('Error adding BOM item:', error);
        showToast('Failed to add BOM item', 'error');
    }
}

function filterBOMList() {
    const filter = document.getElementById('bomSearchInput').value;
    renderBOMList(filter);
}

function loadBOMFromFile() {
    try {
        const fileInput = document.getElementById('bomFileInput');
        const file = fileInput.files[0];

        if (!file) {
            showToast('Please select a file first', 'warning');
            return;
        }

        if (!validateFileSize(file)) {
            return;
        }

        showLoading('Loading inventory items...');
        const reader = new FileReader();

        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(firstSheet);

                // Expected columns: PartNumber, Description, Quantity, Unit
                let loadedCount = 0;
                jsonData.forEach(row => {
                    if (row.PartNumber && row.Description) {
                        const newBOM = {
                            id: 'BOM' + Date.now() + '_' + loadedCount,
                            partNumber: sanitizeInput(row.PartNumber.toString()),
                            description: sanitizeInput(row.Description.toString()),
                            quantity: row.Quantity ? parseInt(row.Quantity) : 1,
                            unit: row.Unit ? sanitizeInput(row.Unit.toString()) : 'pcs'
                        };
                        bomRecords.push(newBOM);
                        loadedCount++;
                    }
                });

                renderBOMList();
                fileInput.value = ''; // Clear file input
                hideLoading();
                showToast(`Loaded ${loadedCount} inventory items successfully`, 'success');
            } catch (error) {
                hideLoading();
                console.error('Error reading BOM file:', error);
                showToast('Failed to read inventory file. Please check the format (PartNumber, Description, Quantity, Unit columns required)', 'error');
            }
        };

        reader.onerror = function() {
            hideLoading();
            showToast('Failed to read file', 'error');
        };

        reader.readAsArrayBuffer(file);
    } catch (error) {
        hideLoading();
        console.error('Error loading BOM file:', error);
        showToast('An error occurred while loading the inventory file', 'error');
    }
}

// Copy BOM assignments from source asset to selected assets
function copyBOMFromAsset() {
    try {
        const sourceNodeId = document.getElementById('bomCopySource').value;
        if (!sourceNodeId) {
            showToast('Please select a source asset to copy from', 'warning');
            return;
        }

        const sourceBOMs = assetBomAssignments[sourceNodeId] || [];
        if (sourceBOMs.length === 0) {
            showToast('Selected asset has no BOM items assigned', 'info');
            return;
        }

        const affectedNodeIds = getAffectedNodes(selectedNodes);
        affectedNodeIds.forEach(nodeId => {
            if (!assetBomAssignments[nodeId]) {
                assetBomAssignments[nodeId] = [];
            }
            // Copy all BOMs from source, avoiding duplicates
            sourceBOMs.forEach(bomId => {
                if (!assetBomAssignments[nodeId].includes(bomId)) {
                    assetBomAssignments[nodeId].push(bomId);
                }
            });
        });

        renderAssignedBOMs();
        renderBOMList(document.getElementById('bomSearchInput').value);
        updateTreeView(); // Refresh tree to show updated badges

        const tree = $('#treeView').jstree(true);
        const sourceNode = tree.get_node(sourceNodeId);
        showToast(`Copied ${sourceBOMs.length} BOM item(s) from "${sourceNode.text}" to ${affectedNodeIds.length} asset(s)`, 'success');
    } catch (error) {
        console.error('Error copying BOM assignments:', error);
        showToast('Failed to copy BOM assignments', 'error');
    }
}

function displaySelectedNodes(containerId) {
    const container = document.getElementById(containerId);
    container.innerHTML = '';

    if (selectedNodes.length === 0) {
        container.innerHTML = '<p class="placeholder-text">No nodes selected</p>';
        return;
    }

    selectedNodes.forEach(node => {
        const tag = document.createElement('div');
        tag.className = 'selected-node-tag';
        tag.innerHTML = `
            <span>${sanitizeInput(node.text)}</span>
            <button class="remove-tag" data-node-id="${node.id}">×</button>
        `;

        const removeBtn = tag.querySelector('.remove-tag');
        removeBtn.addEventListener('click', () => {
            selectedNodes = selectedNodes.filter(n => n.id !== node.id);
            displaySelectedNodes(containerId);
            // Don't auto-close modals when removing nodes
            // User can use the Close button to close the modal
        });

        container.appendChild(tag);
    });
}

// Close modals when clicking outside
window.addEventListener('click', function(event) {
    const pmModal = document.getElementById('pmManagementModal');
    const bomModal = document.getElementById('bomManagementModal');

    if (event.target === pmModal) {
        closePMManagementModal();
    }

    if (event.target === bomModal) {
        closeBOMManagementModal();
    }
});

// Update PM/BOM details section in the edit panel
function updatePMBOMDetails(node) {
    const detailsSection = document.getElementById('pmBomDetailsSection');
    const pmDetailsList = document.getElementById('pmDetailsList');
    const bomDetailsList = document.getElementById('bomDetailsList');
    const pmCountElement = document.getElementById('pmDetailsCount');
    const bomCountElement = document.getElementById('bomDetailsCount');

    if (!node) {
        detailsSection.style.display = 'none';
        return;
    }

    const nodeId = node.id;
    const pmAssignments = assetPmAssignments[nodeId] || [];
    const bomAssignments = assetBomAssignments[nodeId] || [];

    // Show the section if there are assignments
    if (pmAssignments.length > 0 || bomAssignments.length > 0) {
        detailsSection.style.display = 'block';
    } else {
        detailsSection.style.display = 'none';
        return;
    }

    // Update PM count
    pmCountElement.textContent = pmAssignments.length;

    // Populate PM list
    if (pmAssignments.length > 0) {
        pmDetailsList.innerHTML = '';
        pmAssignments.forEach(pmId => {
            const pm = pmRecords.find(p => p.id === pmId);
            if (pm) {
                const item = document.createElement('div');
                item.className = 'details-item pm-item';
                item.innerHTML = `
                    <div class="details-item-code">${sanitizeInput(pm.code)}</div>
                    <div class="details-item-description">${sanitizeInput(pm.description)} | ${sanitizeInput(pm.frequency)}</div>
                `;
                pmDetailsList.appendChild(item);
            }
        });
    } else {
        pmDetailsList.innerHTML = '<p class="placeholder-text">No PM tasks assigned</p>';
    }

    // Update BOM count
    bomCountElement.textContent = bomAssignments.length;

    // Populate BOM list
    if (bomAssignments.length > 0) {
        bomDetailsList.innerHTML = '';
        bomAssignments.forEach(bomId => {
            const bom = bomRecords.find(b => b.id === bomId);
            if (bom) {
                const item = document.createElement('div');
                item.className = 'details-item bom-item';
                item.innerHTML = `
                    <div class="details-item-code">${sanitizeInput(bom.partNumber)}</div>
                    <div class="details-item-description">${sanitizeInput(bom.description)} | Qty: ${bom.quantity} ${bom.unit}</div>
                `;
                bomDetailsList.appendChild(item);
            }
        });
    } else {
        bomDetailsList.innerHTML = '<p class="placeholder-text">No BOM items assigned</p>';
    }
}

// Show details when PM/BOM item node is clicked
function updatePMBOMItemDetails(node) {
    const editForm = document.getElementById('editForm');
    const editPanelTitle = document.getElementById('editPanelTitle');
    const pmBomDetailsSection = document.getElementById('pmBomDetailsSection');

    // Hide the PM/BOM assignments section
    pmBomDetailsSection.style.display = 'none';

    if (node.type === 'pm-item' && node.data && node.data.pm) {
        const pm = node.data.pm;
        editPanelTitle.textContent = 'PM Task Details';
        editForm.innerHTML = `
            <div class="pm-bom-detail-view">
                <div class="detail-header pm-header">
                    <i class="fas fa-clipboard-check"></i>
                    <h4>Preventive Maintenance Task</h4>
                </div>
                <div class="detail-field">
                    <label>PM Code:</label>
                    <div class="detail-value">${sanitizeInput(pm.code)}</div>
                </div>
                <div class="detail-field">
                    <label>Description:</label>
                    <div class="detail-value">${sanitizeInput(pm.description)}</div>
                </div>
                <div class="detail-field">
                    <label>Frequency:</label>
                    <div class="detail-value">${sanitizeInput(pm.frequency)}</div>
                </div>
                <div class="detail-field">
                    <label>Type:</label>
                    <div class="detail-value">${sanitizeInput(pm.type)}</div>
                </div>
                <div class="detail-field">
                    <label>Assigned to Asset:</label>
                    <div class="detail-value">${sanitizeInput(node.data.assetId)}</div>
                </div>
            </div>
        `;
    } else if (node.type === 'bom-item' && node.data && node.data.bom) {
        const bom = node.data.bom;
        editPanelTitle.textContent = 'BOM Item Details';
        editForm.innerHTML = `
            <div class="pm-bom-detail-view">
                <div class="detail-header bom-header">
                    <i class="fas fa-box"></i>
                    <h4>Bill of Materials Item</h4>
                </div>
                <div class="detail-field">
                    <label>Part Number:</label>
                    <div class="detail-value">${sanitizeInput(bom.partNumber)}</div>
                </div>
                <div class="detail-field">
                    <label>Description:</label>
                    <div class="detail-value">${sanitizeInput(bom.description)}</div>
                </div>
                <div class="detail-field">
                    <label>Quantity:</label>
                    <div class="detail-value">${bom.quantity} ${sanitizeInput(bom.unit)}</div>
                </div>
                <div class="detail-field">
                    <label>Assigned to Asset:</label>
                    <div class="detail-value">${sanitizeInput(node.data.assetId)}</div>
                </div>
            </div>
        `;
    }
}

// Show folder summary
function showFolderSummary(node) {
    const editForm = document.getElementById('editForm');
    const editPanelTitle = document.getElementById('editPanelTitle');
    const pmBomDetailsSection = document.getElementById('pmBomDetailsSection');

    // Hide the PM/BOM assignments section
    pmBomDetailsSection.style.display = 'none';

    const childCount = node.children ? node.children.length : 0;

    if (node.type === 'pm-folder') {
        editPanelTitle.textContent = 'PM Tasks Folder';
        editForm.innerHTML = `
            <div class="pm-bom-detail-view">
                <div class="detail-header pm-header">
                    <i class="fas fa-clipboard-list"></i>
                    <h4>PM Tasks Collection</h4>
                </div>
                <div class="detail-field">
                    <label>Total PM Tasks:</label>
                    <div class="detail-value">${childCount}</div>
                </div>
                <p class="placeholder-text">Expand this folder to view individual PM tasks</p>
            </div>
        `;
    } else if (node.type === 'bom-folder') {
        editPanelTitle.textContent = 'BOM Items Folder';
        editForm.innerHTML = `
            <div class="pm-bom-detail-view">
                <div class="detail-header bom-header">
                    <i class="fas fa-boxes"></i>
                    <h4>BOM Items Collection</h4>
                </div>
                <div class="detail-field">
                    <label>Total BOM Items:</label>
                    <div class="detail-value">${childCount}</div>
                </div>
                <p class="placeholder-text">Expand this folder to view individual BOM items</p>
            </div>
        `;
    }
}

document.addEventListener('DOMContentLoaded', function () {
    const dataForm = document.getElementById('dataForm');
    const dataTable = document.getElementById('dataTable').getElementsByTagName('tbody')[0];

    let db;

    // Abrir ou criar banco de dados
    const request = indexedDB.open('UserDataDB', 1);

    request.onerror = function(event) {
        console.error('Erro ao abrir o banco de dados', event);
    };

    request.onsuccess = function(event) {
        db = event.target.result;
        loadDataFromDB();
    };

    request.onupgradeneeded = function(event) {
        db = event.target.result;
        const objectStore = db.createObjectStore('users', { keyPath: 'id', autoIncrement: true });
        objectStore.createIndex('name', 'name', { unique: false });
        objectStore.createIndex('age', 'age', { unique: false });
    };

    dataForm.addEventListener('submit', function (event) {
        event.preventDefault();
        const name = document.getElementById('name').value;
        const age = document.getElementById('age').value;

        addDataToDB({ name, age });
        dataForm.reset();
    });

    function addDataToDB(data) {
        const transaction = db.transaction(['users'], 'readwrite');
        const objectStore = transaction.objectStore('users');
        const request = objectStore.add(data);

        request.onsuccess = function(event) {
            addDataToTable(data, event.target.result);
        };

        request.onerror = function(event) {
            console.error('Erro ao adicionar dados', event);
        };
    }

    function loadDataFromDB() {
        const transaction = db.transaction(['users'], 'readonly');
        const objectStore = transaction.objectStore('users');

        objectStore.openCursor().onsuccess = function(event) {
            const cursor = event.target.result;
            if (cursor) {
                addDataToTable(cursor.value, cursor.key);
                cursor.continue();
            }
        };
    }

    function addDataToTable(data, key) {
        const row = dataTable.insertRow();
        row.setAttribute('data-id', key);
        const cell1 = row.insertCell(0);
        const cell2 = row.insertCell(1);
        const cell3 = row.insertCell(2);

        cell1.textContent = data.name;
        cell2.textContent = data.age;

        const editButton = document.createElement('button');
        editButton.textContent = 'Editar';
        editButton.onclick = function () {
            editData(row, data.name, data.age, key);
        };
        const deleteButton = document.createElement('button');
        deleteButton.textContent = 'Excluir';
        deleteButton.onclick = function () {
            deleteData(row, key);
        };

        cell3.appendChild(editButton);
        cell3.appendChild(deleteButton);
    }

    function editData(row, name, age, key) {
        const newName = prompt("Editar nome:", name);
        const newAge = prompt("Editar idade:", age);
        if (newName && newAge) {
            const transaction = db.transaction(['users'], 'readwrite');
            const objectStore = transaction.objectStore('users');
            const request = objectStore.put({ name: newName, age: newAge, id: key });

            request.onsuccess = function() {
                row.cells[0].textContent = newName;
                row.cells[1].textContent = newAge;
            };

            request.onerror = function(event) {
                console.error('Erro ao editar dados', event);
            };
        }
    }

    function deleteData(row, key) {
        const transaction = db.transaction(['users'], 'readwrite');
        const objectStore = transaction.objectStore('users');
        const request = objectStore.delete(key);

        request.onsuccess = function() {
            dataTable.deleteRow(row.rowIndex - 1);
        };

        request.onerror = function(event) {
            console.error('Erro ao excluir dados', event);
        };
    }

    window.exportToExcel = function() {
        const data = [];
        const transaction = db.transaction(['users'], 'readonly');
        const objectStore = transaction.objectStore('users');

        objectStore.openCursor().onsuccess = function(event) {
            const cursor = event.target.result;
            if (cursor) {
                data.push({ name: cursor.value.name, age: cursor.value.age });
                cursor.continue();
            } else {
                const worksheet = XLSX.utils.json_to_sheet(data);
                const workbook = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
                XLSX.writeFile(workbook, 'meus_dados.xlsx');
            }
        };
    };
});

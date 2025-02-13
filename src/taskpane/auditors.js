Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        // Menambahkan event listener ke tombol
        document.getElementById("advanceLookup").addEventListener("click", showLookupUI);
        document.getElementById("createNewGroup").addEventListener("click", showNewGroupUI);
        document.getElementById("addCategory").addEventListener("click", addCategory);
        document.getElementById("saveGroup").addEventListener("click", saveGroup);
    }
});

function showLookupUI() {
    // Menampilkan UI Lookup
    document.getElementById("lookupUI").style.display = "block";
}

function showNewGroupUI() {
    // Menampilkan UI New Group
    document.getElementById("newGroupUI").style.display = "block";
}

function addCategory() {
    // Logika untuk menambahkan kategori
    alert("Add Category clicked!");
}

function saveGroup() {
    // Logika untuk menyimpan group
    alert("Save Group clicked!");
}
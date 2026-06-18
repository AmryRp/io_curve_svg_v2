// Function to recursively rename items and groups
function renameGroupItems(group, newName, counter) {
    // Rename the group itself
    var oldGroupName = group.name;
    group.name = newName + "_Group_" + counter.value;
    // alert("Renamed group " + oldGroupName + " to " + group.name);
    counter.value++;

    for (var i = 0; i < group.pageItems.length; i++) {
        var item = group.pageItems[i];

        if (item.typename === 'GroupItem') {
            // alert("Entering subgroup: " + item.name);
            renameGroupItems(item, newName, counter); // Recursively rename items in nested groups
            // alert("Exiting subgroup: " + item.name);
        } else {
            var oldName = item.name;
            item.name = newName + "_" + counter.value;
            // alert("Renamed " + oldName + " to " + item.name);
            counter.value++;
        }
    }
}

// Main script
function main() {
    // Get the active document
    var doc = app.activeDocument;

    // Check if there's an active selection
    if (doc.selection.length > 0) {
        // Prompt the user to enter the new name
        var newName = prompt("Enter the new name for the items in the group:", "NewName");

        if (newName !== null) {
            // Initialize a counter object to keep track of unique names
            var counter = { value: 1 };

            // Loop through selected items
            for (var i = 0; i < doc.selection.length; i++) {
                var selectedItem = doc.selection[i];

                // Check if the selected item is a group
                if (selectedItem.typename === 'GroupItem') {
                    // alert("Renaming items in group: " + selectedItem.name);
                    renameGroupItems(selectedItem, newName, counter);
                    // alert("Finished renaming items in group: " + selectedItem.name);
                } else {
                    alert("Please select a group to rename its items.");
                }
            }
        }
    } else {
        alert("Please select a group to rename its items.");
    }
}

// Run the main script
main();

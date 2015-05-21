var listitemid;
var listitemidDict;
var currentWeb;
var currentSite;
var currentListGuid;

function Export() {
    
    var selecteditems = SP.ListOperation.Selection.getSelectedItems();
    currentListGuid = SP.ListOperation.Selection.getSelectedList();
    var context = SP.ClientContext.get_current();
    currentSite = context.get_site();

    currentWeb = context.get_web();
    var currentList = currentWeb.get_lists().getById(currentListGuid);

    var index;
    listitemidDict = '';

    for (index in selecteditems) {
        listitemid = currentList.getItemById(selecteditems[index].id);
        listitemidDict = listitemidDict + selecteditems[index].id + ',';        
    }

    context.executeQueryAsync(Function.createDelegate(this, this.success), Function.createDelegate(this, this.failed));
}

function success() {
    
    var form = document.createElement("form");
    form.setAttribute("method", "post");   

    var hiddenField = document.createElement("input");
    hiddenField.setAttribute("type", "hidden");
    hiddenField.setAttribute("name", "IDDict");
    hiddenField.setAttribute("value", listitemidDict);
    form.appendChild(hiddenField);

    var hiddenListGuid = document.createElement("input");
    hiddenListGuid.setAttribute("type", "hidden");
    hiddenListGuid.setAttribute("name", "ListGuid");
    hiddenListGuid.setAttribute("value", currentListGuid);
    form.appendChild(hiddenListGuid); 

    var hiddenViewGuid = document.createElement("input");
    hiddenViewGuid.setAttribute("type", "hidden");
    hiddenViewGuid.setAttribute("name", "ViewGuid");
    hiddenViewGuid.setAttribute("value", ctx.view);
    form.appendChild(hiddenViewGuid);

    form.setAttribute("action", ctx.HttpRoot+"/_layouts/Rjesh.Solutions/ExportToExcel.aspx");    
    document.body.appendChild(form);
    form.submit();    

    SP.UI.Notify.addNotification('Exported Successfully');
}

function failed(sender, args) {
    var statusId = SP.UI.Status.addStatus(args.get_message());
    SP.UI.Status.setStatusPriColor(statusId, 'red');
    latestId = statusId;
}

function exporttoexcelenable() {
    return (true);
}
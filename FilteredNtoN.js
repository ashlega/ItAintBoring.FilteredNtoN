// Custom function to define filters per relationship
function getAddExistingFilters(relationshipName, primaryEntityName)
{
	if(relationshipName == "ita__ita__finding_ita__complaint" && primaryEntityName == "ita__complaint")
	{
		return [{entityLogicalName: "ita__finding", filterXml: "<filter type='and'><condition attribute='ita__mainentity' operator='eq' value='" + Xrm.Page.getAttribute("ita__mainentity").getValue()[0].id + "' /></filter>"}]
	}
	if(relationshipName == "ita__ita__finding_ita__complaint" && primaryEntityName == "ita__finding")
	{
		return [{entityLogicalName: "ita__complaint", filterXml: "<filter type='and'><condition attribute='ita__mainentity' operator='eq' value='" + Xrm.Page.getAttribute("ita__mainentity").getValue()[0].id + "' /></filter>"}]
	}
	return null;
}


// Custom function to call instead of the OOTB Add Existing button/command - all 4 parameters can be passed from the ribbon
function filterAddExisting(selectedEntityTypeName, selectedControl, firstPrimaryItemId, relationshipList) {
	var relationshipName = selectedControl.getRelationship().name;
	var primaryEntityName = Xrm.Page.data.entity.getEntityName();
	if (relationshipList.indexOf(relationshipName) > -1) {
        var options = {
            allowMultiSelect: true,
            entityTypes: [selectedEntityTypeName],
            showNew: true,
			disableMru: true,
            filters: getAddExistingFilters(relationshipName, primaryEntityName)
        };

        lookupAddExistingRecords(relationshipName, selectedEntityTypeName, primaryEntityName, firstPrimaryItemId, selectedControl, options);
    }
    else {
        // Any other contact relationship (N:N or 1:N) - use default behaviour
        XrmCore.Commands.AddFromSubGrid.addExistingFromSubGridAssociated(selectedEntityTypeName, selectedControl);
    }
}

// relationshipName = the schema name of the N:N or 1:N relationship
// primaryEntity = the 1 in the 1:N or the first entity in the N:N - for N:N this is the entity which was used to create the N:N (may need to trial and error this)
// relatedEntity = the N in the 1:N or the secondary entity in the N:N
// parentRecordId = the guid of the record this subgrid/related entity is used on
// gridControl = the grid control parameter passed from the ribbon context
// lookupOptions = options for creating the custom lookup with filters: http://butenko.pro/2017/11/22/microsoft-dynamics-365-v9-0-lookupobjects-closer-look/
function lookupAddExistingRecords(relationshipName, primaryEntity, relatedEntity, parentRecordId, gridControl, lookupOptions) {
    Xrm.Utility.lookupObjects(lookupOptions).then(function (results) {
        // Get the entitySet name for the primary entity
        Xrm.Utility.getEntityMetadata(primaryEntity).then(function (primaryEntityData) {
            var primaryEntitySetName = primaryEntityData.EntitySetName;

            // Get the entitySet name for the related entity
            Xrm.Utility.getEntityMetadata(relatedEntity).then(function (relatedEntityData) {
                var relatedEntitySetName = relatedEntityData.EntitySetName;

                // Call the associate web api for each result (recursive)
                associateAddExistingResults(relationshipName, primaryEntitySetName, relatedEntitySetName, relatedEntity, parentRecordId.replace("{", "").replace("}", ""), gridControl, results, 0)
            });
        });
    });
}

// Used internally by the above function
function associateAddExistingResults(relationshipName, primaryEntitySetName, relatedEntitySetName, relatedEntity, parentRecordId, gridControl, results, index) {
    if (index >= results.length) {
        // Refresh the grid once completed
        Xrm.Page.ui.setFormNotification("Associated " + index + " record" + (index > 1 ? "s" : ""), "INFO", "associate");
        if (gridControl) { gridControl.refresh(); }

        // Clear the final notification after 2 seconds
        setTimeout(function () {
            Xrm.Page.ui.clearFormNotification("associate");
        }, 2000);

        return;
    }

    Xrm.Page.ui.setFormNotification("Associating record " + (index + 1) + " of " + results.length, "INFO", "associate");

    var lookupId = results[index].id.replace("{", "").replace("}", "");
    var lookupEntity = results[index].entityType || results[index].typename;

    var primaryId = parentRecordId;
    var relatedId = lookupId;
    if (lookupEntity.toLowerCase() != relatedEntity.toLowerCase()) {
        // If the related entity is different to the lookup entity flip the primary and related id's
        primaryId = lookupId;
        relatedId = parentRecordId;
    }

    var association = { '@odata.id': Xrm.Page.context.getClientUrl() + "/api/data/v9.0/" + relatedEntitySetName + "(" + relatedId + ")" };

    var req = new XMLHttpRequest();
    req.open("POST", Xrm.Page.context.getClientUrl() + "/api/data/v9.0/" + primaryEntitySetName + "(" + primaryId + ")/" + relationshipName + "/$ref", true);
    req.setRequestHeader("Accept", "application/json");
    req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    req.setRequestHeader("OData-MaxVersion", "4.0");
    req.setRequestHeader("OData-Version", "4.0");
    req.onreadystatechange = function () {
        if (this.readyState === 4) {
            req.onreadystatechange = null;
            index++;
            if (this.status === 204 || this.status === 1223) {
                // Success
                // Process the next item in the list
                associateAddExistingResults(relationshipName, primaryEntitySetName, relatedEntitySetName, relatedEntity, parentRecordId, gridControl, results, index);
            }
            else {
                // Error
                var error = JSON.parse(this.response).error.message;
                if (error == "A record with matching key values already exists.") {
                    // Process the next item in the list
                    associateAddExistingResults(relationshipName, primaryEntitySetName, relatedEntitySetName, relatedEntity, parentRecordId, gridControl, results, index);
                }
                else {
                    Xrm.Utility.alertDialog(error);
                    Xrm.Page.ui.clearFormNotification("associate");
                    if (gridControl) { gridControl.refresh(); }
                }
            }
        }
    };
    req.send(JSON.stringify(association));
}
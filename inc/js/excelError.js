//this object contains messages that are specific to the error received by the response.
//this allows use to react to a specific error in a specific way rather than just have a readable message that is create on the server end.
var excelError = { 
                    errorType : 
                    {
                        general : 
                        {
                            message: "An error has been encountered"
                        },
                        mergedcells : 
                        {
                            message : "Your document contains merged cells. Please unmerged the cells and try again."
                        }
                    }
                 };

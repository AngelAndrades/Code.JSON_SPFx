import { sp } from '@pnp/sp/presets/all';
import { map, distinct, toArray } from 'rxjs/operators';
import { from } from 'rxjs';
import { Store } from './../state/store';
import * as $ from 'jquery';
import '@progress/kendo-ui';

export class SPA {
    private static instance: SPA;

    private constructor() {}

    public static getInstance(store: Store): SPA {
        let tabStrip: kendo.ui.TabStrip = null;
        let importGrid: kendo.ui.Grid = null;
        let appendGrid: kendo.ui.Grid = null;
        let filter: kendo.ui.Filter = null;
        let dialog: kendo.ui.Dialog = null;
        let progressBar: kendo.ui.ProgressBar = null;

        const cleanseItem = (dataItem: object): object => {
            // remove unwanted properties before saving
            $.each(dataItem, (k,v) => {
                switch (k) {
                    case 'Title':
                    case 'codeVersion':
                    case 'usageType':
                    case 'licenseName':
                    case 'opRL':
                    case 'downloadURL':
                    case 'homepageURL':
                    case 'disclaimerURL':
                    case 'repositoryURL':
                    case 'vcs':
                    case 'laborHours':
                    case 'tags':
                        break;
                    default:
                        delete dataItem[k];
                }
            });

            return dataItem;
        };

        $(() => {
            const tabStripOptions: kendo.ui.TabStripOptions = {
                select: (e:kendo.ui.TabStripSelectEvent) => {
                    if (e.item.textContent === 'User Guide') {
                        e.preventDefault();
                        tabStrip.select(0);
                        window.open(store.value.spLink);
                    }
                }
            };

            tabStrip = $('#tabStrip').kendoTabStrip(tabStripOptions).data('kendoTabStrip');

            const importGridOptions: kendo.ui.GridOptions = {
                dataSource: new kendo.data.DataSource({
                    transport: {
                        read: async options => {
                            await sp.web.lists.getById(store.value.importList).items.top(1000).getPaged()
                            .then(response => {
                                const recurse = (next: any) => {
                                    next.getNext().then(nestedResponse => {
                                        store.set('dsImportData', [...store.value.dsImportData, ...nestedResponse.results]);
                                        if (nestedResponse.hasNext) recurse(nestedResponse);
                                        else options.success(store.value.dsImportData);
                                    });
                                };

                                store.set('dsImportData', response.results);
                                if (response.hasNext) recurse(response);
                                else options.success(store.value.dsImportData);
                            })
                            .catch(error => {
                                console.log(error);
                                throw new Error('ImportData error, unable to read items');
                            });
                        }
                    }
                }),
                height: 550,
                pageable: true,
                sortable: true,
                toolbar: [ 'search' ],
                columns: [
                    { field: 'VASI_x0020_Id', title: 'VASI Id', width: 100 },
                    { field: 'System_x0020_Name', title: 'System Name', width: 250 },
                    { field: 'Software_x0020_Type', title: 'Software Type', width: 150 },
                    { field: 'System_x0020_Status', title: 'System Status', width: 150 },
                    { field: 'System_x0020_Description', title: 'Description', hidden: true }
                ]
            };

            importGrid = $('#importGrid').kendoGrid(importGridOptions).data('kendoGrid');

            const Editors = {
                onSystemStatus: (container, options) => {
                    let ds1 = from(importGrid.dataSource.view()).pipe(
                        map(o => o.System_x0020_Status),
                        distinct(),
                        toArray()
                    ).subscribe(out => {
                        $('<input data-bind="value: value" name="' + options.field + '"/>')
                        .appendTo(container)
                        .kendoDropDownList({
                            dataSource: out
                        });
                    });
                },
                onSoftwareType: (container, options) => {
                    let ds2 = from (importGrid.dataSource.view()).pipe(
                        map(o => o.Software_x0020_Type),
                        distinct(),
                        toArray()
                    ).subscribe(out => {
                        $('<input data-bind="value: value" name="' + options.field + '"/>')
                        .appendTo(container)
                        .kendoDropDownList({
                            dataSource: out.filter(o => o != null)
                        });
                    });
                },
                onVasiId: (container, options) => {
                    $('<input required name="' + options.field + '"/>')
                    .appendTo(container)
                    .kendoMultiColumnComboBox({
                        autoBind: true,
                        dataTextField: "System_x0020_Name",
                        dataValueField: "VASI_x0020_Id",
                        dataSource: { data: importGrid.dataSource.view() },
                        columns: [
                            { field: 'VASI_x0020_Id', title: 'VASI ID', width: 100 },
                            { field: 'System_x0020_Name', title: 'System Name', width: 300 }
                        ]
                    });
                },
                onUsageType: (container, options) => {
                    $('<input required name="' + options.field + '"/>')
                    .appendTo(container)
                    .kendoDropDownList({
                        dataSource: ['openSource','governmentWideReuse']
                    });
                },
                onLicenseName: (container, options) => {
                    $('<input name="' + options.field + '"/>')
                    .appendTo(container)
                    .kendoDropDownList({
                        dataSource: [' ', 'Public Domain, CC0-1.0', 'Permissive', 'LGPL', 'Copyleft', 'Proprietary']
                    });
                }
            };

            const filterOptions: kendo.ui.FilterOptions = {
                dataSource: importGrid.dataSource,
                expressionPreview: false,
                applyButton: true,
                fields: [
                    { name: 'System_x0020_Status', type: 'string', label: 'System Status', editorTemplate: Editors.onSystemStatus },
                    { name: 'Software_x0020_Type', type: 'string', label: 'Software Type', editorTemplate: Editors.onSoftwareType }
                ],
                expression: {
                    logic: 'and',
                    filters: [
                        { field: 'Software_x0020_Type', value: 'Custom Development', operator: 'eq' },
                        { logic: 'or', filters: [
                            { field: 'System_x0020_Status', value: 'Development', operator: 'eq' },
                            { field: 'System_x0020_Status', value: 'Production', operator: 'eq' },
                            { field: 'System_x0020_Status', value: 'Inactive', operator: 'eq' }
                        ]}
                    ]
                }
            };

            filter = $('#filter').kendoFilter(filterOptions).data('kendoFilter');

            const Templates = {
                vasiId: dataItem => {
                    let itemRef = importGrid.dataSource.view().find((e:any) => parseInt(e.VASI_x0020_Id) === parseInt(dataItem.Title));
                    let vasiIdTitle = (Object(itemRef).System_x0020_Name != undefined) ? dataItem.Title + '<br/>' + Object(itemRef).System_x0020_Name : dataItem.Title;
                    return dataItem.Id !== '' ? vasiIdTitle : dataItem.Title;
                }
            };

            const appendGridOptions: kendo.ui.GridOptions = {
                dataSource: new kendo.data.DataSource({
                    transport: {
                        create: async options => {
                            await sp.web.lists.getById(store.value.appendList).items.add(cleanseItem(options.data))
                            .then(response => {
                                options.success(response.data);
                                $('#appendGrid').data('kendoGrid').dataSource.read();  //reload data so system name can be added dynamically
                            })
                            .catch(error => {
                                console.log(error);
                                throw new Error('AppendData error, unable to create item');
                            });
                        },
                        read: async options => {
                            await sp.web.lists.getById(store.value.appendList).items.select('Id','Title','codeVersion','laborHours','licenseName','opRL','repositoryURL','tags','usageType','Created','Modified').top(1000).getPaged()
                            .then(response => {
                                const recurse = (next: any) => {
                                    next.getNext().then(nestedResponse => {
                                        $.each(nestedResponse.results, (index,value) => {
                                            let dataItem = importGrid.dataSource.data().find((o:any) => parseInt(o.VASI_x0020_Id) == parseInt(value.Title));
                                            value.systemName = dataItem.System_x0020_Name;
                                        });
                                        store.set('dsAppendData', [...store.value.dsAppendData, ...nestedResponse.results]);
                                        if (nestedResponse.hasNext) recurse(nestedResponse);
                                        else options.success(store.value.dsAppendData);
                                    });
                                };

                                // delay loading the data, otherwise importgrid datasource will not be ready
                                setTimeout(() => {
                                    $.each(response.results, (index,value) => {
                                        let dataItem = importGrid.dataSource.data().find((o:any) => parseInt(o.VASI_x0020_Id) == parseInt(value.Title));
                                        value.systemName = dataItem.System_x0020_Name;
                                    });
                                    store.set('dsAppendData', response.results);

                                    if (response.hasNext) recurse(response);
                                    else options.success(store.value.dsAppendData);
                                }, 2000);
                            })
                            .catch(error => {
                                console.log(error);
                                throw new Error('AppendData error, unable to read items');
                            });
                        },
                        update: async options => {
                            await sp.web.lists.getById(store.value.appendList).items.getById(options.data.Id).update(cleanseItem(options.data))
                            .then(response => {
                                options.success();
                            })
                            .catch(error => {
                                console.log(error);
                                throw new Error('AppendData error, unable to update item');
                            });
                        },
                        destroy: async options => {
                            await sp.web.lists.getById(store.value.appendList).items.getById(options.data.Id).recycle()
                            .then(response => {
                                options.success();
                            })
                            .catch(error => {
                                console.log(error);
                                throw new Error('AppendData error, unable to delete item');
                            });
                        }
                    },
                    pageSize: store.value.pageSize,
                    sort: { field: 'systemName', dir: 'asc' },
                    schema: {
                        model: {
                            id: 'Id',
                            fields: {
                                Title: { type: 'string' },
                                systemName: { type: 'string' },
                                codeVersion: { type: 'string' },
                                usageType: { type: 'string', defaultValue: 'governmentWideResuse' },
                                licenseName: { type: 'string' },
                                opRL: { type: 'string' },
                                repositoryURL: { type: 'string' },
                                vcs: { type: 'string', defaultValue: 'GitHub' },
                                laborHours: { type: 'number', defaultValue: 0 },
                                tags: { type: 'string' },
                                Created: { type: 'date' },
                                Modified: { type: 'date' },
                                metadataLastUpdated: { type: 'date' }
                            }
                        }
                    }
                }),
                columnMenu: true,
                editable: 'popup',
                filterable: false,
                pageable: { 
                    pageSize: 5,
                    pageSizes: true,
                    buttonCount: 4
                },
                scrollable: { virtual: 'columns' },
                sortable: true,
                toolbar: [ 'create', { name: 'export', text: 'Download Code.JSON File', iconClass: 'k-icon k-i-download' }, 'search' ],
                columns: [
                    { command: ['edit', 'destroy'], width: 225 },
                    { field: 'Title', title: 'VASI ID', editor: Editors.onVasiId, width: 100 },
                    { field: 'systemName', title: 'System Name', width: 300 },
                    { field: 'codeVersion', title: 'Software Version', width: 150 },
                    { field: 'tags', title: 'Tags', width: 150 },
                    { field: 'laborHours', title: 'Labor Hours', width: 150 },
                    { field: 'repositoryURL', title: 'Repository URL', width: 1000 },
                    { title: 'Permissions',
                        columns: [
                            { field: 'usageType', title: 'Usage Type', editor: Editors.onUsageType, width: 200 },
                            { field: 'licenseName', title: 'License', editor: Editors.onLicenseName, width: 200 },
                            { field: 'opRL', title: 'License URL', width: 300 }
                        ]
                    }
                ],
                edit: e => {
                    $('[for="systemName"]').parent().next().remove();
                    $('[for="systemName"]').parent().remove();
                }
            };

            appendGrid = $('#appendGrid').kendoGrid(appendGridOptions).data('kendoGrid');

            store.select('dsImportData').subscribe(obs => {
                filter.setOptions({ dataSource: importGrid.dataSource });
                filter.applyFilter();

                appendGrid.refresh();
            });

            const dialogOptions: kendo.ui.DialogOptions = {
                width: '500px',
                title: 'Generating Code.JSON File',
                closable: false,
                modal: false,
                content: '<p>Please wait while the application processes your request. It is merging your appended information with the VASI Extract and connecting with VA\'s GitHub repository to obtain description information for each entry.</p><h3>Progress Status</h3><div id="progressBar"></div>',
                actions: [
                    { text: 'Close when completed' }
                ]
            };

            const progressBarOptions: kendo.ui.ProgressBarOptions = {
                animation: { duration: 100 },
                min: 0,
                max: 0,
                type: 'value',
                complete: e => {
                    dialog.close();  // close the dialog box automatically when progress bar reaches max
                    progressBar = null;
                }
            };

            $('.k-grid-export').click(() => {
                //const readToken = '1253bd4be30747c1dc7b56c7e40ee9dc856d1d21';   // enterprise github account
                const readToken = '20cc9245594ebbe184defd41dc74115969b99b8a'; // personal github account
                const headers = { 
                    'Authorization' : 'token ' + readToken
                };
                let repoDescription = null;
                let counter = 1;
                
                dialog = $('#dialog').kendoDialog(dialogOptions).data('kendoDialog');
                progressBar = $('#progressBar').kendoProgressBar(progressBarOptions).data('kendoProgressBar');
                progressBar.setOptions({ max: appendGrid.dataSource.data().length });

                $.when( dialog.open() ).then(_ => {
                    var jsonResponse = new Object();
                    jsonResponse['agency'] = 'VA';
                    jsonResponse['version'] = '2.0.0';
                    jsonResponse['measurementType'] = { method: 'modules' };
                    jsonResponse['releases'] = [];
                    
                    $.each(importGrid.dataSource.data(), (index,value) => {
                        let dataItem = appendGrid.dataSource.data().find((o:any) => parseInt(o.Title) == parseInt(value.VASI_x0020_Id));
                        if(dataItem != undefined) {
                            if (dataItem.repositoryURL !== null) {
                                progressBar.value(counter++);

                                let repo = new URL(dataItem.repositoryURL);
                                let repoApi = new URL('https://api.github.com/repos' + repo.pathname);
                                
                                $.ajax({
                                    async: false,
                                    url: repoApi.toString(),
                                    method: 'GET',
                                    headers: headers
                                })
                                .done(data => {
                                    repoDescription = data.description;
                                })
                                .fail(xhr => {
                                    repoDescription = 'Repository containing the FOIA Releases for ' + value.System_x0020_Name;
                                });
                            }
    
                            jsonResponse['releases'].push({
                                id: value.VASI_x0020_Id,
                                name: value.System_x0020_Name,
                                organization: store.value.organization,
                                version: dataItem.codeVersion,
                                status: (value.System_x0020_Status == 'Inactive') ? 'Archival' : value.System_x0020_Status,
                                permissions: { usageType: dataItem.usageType, licenses: [{ name: (dataItem.licenseName !== null) ? dataItem.licenseName : '', opRL: (dataItem.opRL !== null) ? dataItem.opRL : '' }]},
                                homepageURL: store.value.homeLink,
                                repositoryURL: (dataItem.repositoryURL !== null) ? dataItem.repositoryURL : '',
                                vcs: (store.value.vcs).toLowerCase(),
                                laborHours: dataItem.laborHours,
                                tags: [ (dataItem.tags !== null) ? dataItem.tags : '' ],
                                languages: (value.Technology_x0020_Components !== null) ? (value.Technology_x0020_Components).split('<br>') : [],
                                contact: { name: store.value.contactName, email: store.value.contactEmail },
                                date: { created: dataItem.Created, lastModified: dataItem.Modified, metadataLastUpdated: dataItem.Modified },
                                description: repoDescription
                            });
                        }
                    });
                    
                    // set delay to allow progress animation to function correctly, 100 milliseconds * number of records
                    setTimeout(_ => {
                        var a = document.createElement('a');
                        var file = new Blob([JSON.stringify(jsonResponse)], {type: 'application/json'});
                        a.href = URL.createObjectURL(file);
                        a.download = 'code.JSON';
                        a.click();
                    }, 100 * appendGrid.dataSource.data().length);
                });
    
            });

        });

        return SPA.instance;
    }
}
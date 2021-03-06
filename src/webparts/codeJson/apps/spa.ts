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
                    case 'disclaimer':
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
                        dataSource: [' ', 'Creative Commons Zero v1.0 Universal', 'Permissive', 'LGPL', 'Copyleft', 'Proprietary']
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
                            await sp.web.lists.getById(store.value.appendList).items.select('Id','Title','codeVersion','laborHours','licenseName','opRL','repositoryURL','tags','usageType','Created','Modified','disclaimer').top(1000).getPaged()
                            .then(response => {
                                const recurse = (next: any) => {
                                    next.getNext().then(nestedResponse => {
                                        $.each(nestedResponse.results, (index,value) => {
                                            let dataItem = importGrid.dataSource.data().find((o:any) => parseInt(o.VASI_x0020_Id) == parseInt(value.Title));
                                            value.systemName = (dataItem != null) ? dataItem.System_x0020_Name : '';
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
                                        value.systemName = (dataItem != null) ? dataItem.System_x0020_Name : '';
                                    });
                                    store.set('dsAppendData', response.results);

                                    if (response.hasNext) recurse(response);
                                    else options.success(store.value.dsAppendData);
                                }, 3000);
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
                                disclaimer: { type: 'string' },
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
                            { field: 'licenseName', title: 'License', width: 200 },
                            { field: 'opRL', title: 'License URL', width: 300 }
                        ]
                    },
                    { field: 'disclaimer', title: 'Disclaimer', width: 300 }
                ],
                edit: e => {
                    $('[for="systemName"]').parent().next().remove();
                    $('[for="systemName"]').parent().remove();
                },
                save: e => {
                    let dataItem = importGrid.dataSource.data().find((o:any) => parseInt(o.VASI_x0020_Id) == parseInt(e.model['Title']));
                    if (dataItem == null) {
                        alert('The VASI ID entered could not be validated against the current VASI Extract. Please update the VASI Extract List or use the dropdown to locate the appropriate system record.');
                    }
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
                let counter = 1;
                
                dialog = $('#dialog').kendoDialog(dialogOptions).data('kendoDialog');
                progressBar = $('#progressBar').kendoProgressBar(progressBarOptions).data('kendoProgressBar');
                progressBar.setOptions({ max: appendGrid.dataSource.data().length });

                var jsonResponse = new Object();
                jsonResponse['agency'] = 'VA';
                jsonResponse['version'] = '2.0.0';
                jsonResponse['measurementType'] = { method: 'modules' };
                jsonResponse['releases'] = [];
                
                (async () => {
                    for await (let obj of importGrid.dataSource.data().slice(0)) {
                        let dataItem = appendGrid.dataSource.data().find((o:any) => parseInt(o.Title) == parseInt(obj.VASI_x0020_Id));
                        if(dataItem != undefined) {
                            if (dataItem.repositoryURL !== null) {
                                progressBar.value(counter++);

                                let customTags = [];
                                if (dataItem.tags !== null) customTags.push(dataItem.tags);
                                customTags.push((obj.System_x0020_Status == 'Inactive') ? 'Archival' : obj.System_x0020_Status);
                                customTags.push(dataItem.usageType);

                                let customLanguages = [];
                                if (obj.Technology_x0020_Components !== null && (obj.Technology_x0020_Components).indexOf(';') > -1) customLanguages = (obj.Technology_x0020_Components).split(';');
                                if (obj.Technology_x0020_Components !== null && (obj.Technology_x0020_Components).indexOf('<br>') > -1) customLanguages = (obj.Technology_x0020_Components).split('<br>');
                                if (customLanguages.length == 0) customLanguages.push('Information not available.');

                                let record = {
                                    id: obj.VASI_x0020_Id,
                                    name: obj.System_x0020_Name,
                                    organization: store.value.organization,
                                    version: dataItem.codeVersion,
                                    status: (obj.System_x0020_Status == 'Inactive') ? 'Archival' : obj.System_x0020_Status,
                                    permissions: { 
                                        usageType: dataItem.usageType, 
                                        licenses: [{ 
                                            name: (dataItem.licenseName !== null) ? dataItem.licenseName : (store.value.licensing) ? store.value.defaultLicense : '', 
                                            opRL: (dataItem.opRL !== null) ? dataItem.opRL : (store.value.licensing) ? store.value.defaultLicenseUrl : ''
                                        }]
                                    },
                                    homepageURL: store.value.homeLink,
                                    downloadURL: store.value.homeLink,
                                    repositoryURL: (dataItem.repositoryURL !== null) ? dataItem.repositoryURL : '',
                                    disclaimerText: (dataItem.disclaimer != null) ? dataItem.disclaimer : store.value.disclaimer,
                                    vcs: (store.value.vcs).toLowerCase(),
                                    laborHours: parseInt(dataItem.laborHours) + 1,
                                    tags: customTags,
                                    languages: customLanguages,
                                    contact: { 
                                        name: store.value.contactName, 
                                        email: store.value.contactEmail,
                                        URL: 'https://www.va.gov',
                                        phone: '844-698-2311'
                                    },
                                    date: { 
                                        created: kendo.toString(kendo.parseDate(dataItem.Created),'yyyy-MM-dd'),
                                        lastModified: kendo.toString(kendo.parseDate(dataItem.Modified),'yyyy-MM-dd'),
                                        metadataLastUpdated: kendo.toString(kendo.parseDate(dataItem.Modified),'yyyy-MM-dd')
                                    }
                                };

                                if((dataItem.repositoryURL).indexOf('Patches') == -1) {
                                    let repo = new URL(dataItem.repositoryURL);
                                    let repoApi = new URL('https://api.github.com/repos' + repo.pathname);
    
                                    await fetch(repoApi.toString(), {
                                        method: 'GET',
                                        headers: headers
                                    })
                                    .then(response => response.json())
                                    .then(data => {
                                        record['description'] = data['description'];
                                    })
                                    .catch(error => console.log('err: ', error))
                                    .then(() => {
                                        if (record['description'] == undefined) record['description'] = 'Repository containing the FOIA Releases for ' + obj.System_x0020_Name;
                                        jsonResponse['releases'].push(record);
                                    });
                                } else {
                                    if (record['description'] == undefined) record['description'] = 'Repository containing the FOIA Releases for ' + obj.System_x0020_Name;
                                    jsonResponse['releases'].push(record);
                                }
                            }
                        }
                    }
                })().then(() => {
                    var a = document.createElement('a');
                    var file = new Blob([JSON.stringify(jsonResponse)], {type: 'application/json'});
                    a.href = URL.createObjectURL(file);
                    a.download = 'code.JSON';
                    setTimeout(() => { }, 4000);
                    a.click();
                });
            });

        });

        return SPA.instance;
    }
}
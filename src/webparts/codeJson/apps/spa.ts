import { sp } from '@pnp/sp/presets/all';
import { map, distinct, toArray } from 'rxjs/operators';
import { from } from 'rxjs';
import { Store } from './../state/store';
import * as $ from 'jquery';
import '@progress/kendo-ui';

export class SPA {
    private static instance: SPA;

    private constructor() {}

    public static getInstance(extract: string, append: string, link: string): SPA {
        let tabStrip: kendo.ui.TabStrip = null;
        let importGrid: kendo.ui.Grid = null;
        let appendGrid: kendo.ui.Grid = null;
        let filter: kendo.ui.Filter = null;
        let store = new Store();

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
                        window.open(link);
                    }
                }
            };

            tabStrip = $('#tabStrip').kendoTabStrip(tabStripOptions).data('kendoTabStrip');

            const importGridOptions: kendo.ui.GridOptions = {
                dataSource: new kendo.data.DataSource({
                    transport: {
                        read: async options => {
                            await sp.web.lists.getById(extract).items.top(1000).getPaged()
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
                    $('<input required name="' + options.field + '"/>')
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
                            await sp.web.lists.getById(append).items.add(cleanseItem(options.data))
                            .then(response => {
                                options.success(response.data);
                            })
                            .catch(error => {
                                console.log(error);
                                throw new Error('AppendData error, unable to create item');
                            });
                        },
                        read: async options => {
                            await sp.web.lists.getById(append).items.select('Id','Title','codeVersion','disclaimerURL','downloadURL','homepageURL','laborHours','licenseName','opRL','repositoryURL','tags','usageType','vcs','Created','Modified').top(2).getPaged()
                            .then(response => {
                                const recurse = (next: any) => {
                                    next.getNext().then(nestedResponse => {
                                        store.set('dsAppendData', [...store.value.dsAppendData, ...nestedResponse.results]);
                                        if (nestedResponse.hasNext) recurse(nestedResponse);
                                        else options.success(store.value.dsAppendData);
                                    });
                                };

                                store.set('dsAppendData', response.results);
                                if (response.hasNext) recurse(response);
                                else options.success(store.value.dsAppendData);
                            })
                            .catch(error => {
                                console.log(error);
                                throw new Error('AppendData error, unable to read items');
                            });
                        },
                        update: async options => {
                            await sp.web.lists.getById(append).items.getById(options.data.Id).update(cleanseItem(options.data))
                            .then(response => {
                                options.success();
                            })
                            .catch(error => {
                                console.log(error);
                                throw new Error('AppendData error, unable to update item');
                            });
                        },
                        destroy: async options => {
                            await sp.web.lists.getById(append).items.getById(options.data.Id).recycle()
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
                    schema: {
                        model: {
                            id: 'Id',
                            fields: {
                                Title: { type: 'string' },
                                codeVersion: { type: 'string' },
                                usageType: { type: 'string', defaultValue: 'governmentWideResuse' },
                                licenseName: { type: 'string' },
                                opRL: { type: 'string' },
                                homepageURL: { type: 'string' },
                                downloadURL: { type: 'string' },
                                disclaimerURL: { type: 'string' },
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
                pageable: { pageSizes: true },
                scrollable: { virtual: 'columns' },
                sortable: true,
                toolbar: [ 'create', { name: 'export', text: 'Download Code.JSON File', iconClass: 'k-icon k-i-download' }, 'search' ],
                columns: [
                    { command: ['edit', 'destroy'], width: 225 },
                    { field: 'Title', title: 'VASI ID', editor: Editors.onVasiId, template: Templates.vasiId, width: 300 },
                    { field: 'codeVersion', title: 'Software Version', width: 150 },
                    { field: 'tags', title: 'Tags', width: 150 },
                    { field: 'laborHours', title: 'Labor Hours', width: 150 },
                    { title: 'Permissions',
                        columns: [
                            { field: 'usageType', title: 'Usage Type', editor: Editors.onUsageType, width: 200 },
                            { field: 'licenseName', title: 'License', editor: Editors.onLicenseName, width: 200 },
                            { field: 'opRL', title: 'License URL', width: 300 }
                        ]
                    },
                    { title: 'Repository Information', 
                        columns: [
                            { field: 'vcs', title: 'Version Control System', width: 200 },
                            { field: 'homepageURL', title: 'GitHub Homepage URL', width: 300 },
                            { field: 'downloadURL', title: 'Download URL', width: 300 },
                            { field: 'disclaimerURL', title: 'Disclaimer URL', width: 300 },
                            { field: 'repositoryURL', title: 'Repository URL', width: 300 }
                        ]
                    }
                ]
            };

            appendGrid = $('#appendGrid').kendoGrid(appendGridOptions).data('kendoGrid');

            store.select('dsImportData').subscribe(obs => {
                filter.setOptions({ dataSource: importGrid.dataSource });
                filter.applyFilter();

                appendGrid.refresh();
            });

            $('.k-grid-export').click(() => {
                var jsonResponse = new Object();
                jsonResponse['agency'] = 'VA';
                jsonResponse['version'] = '2.0.0';
                jsonResponse['measurementType'] = { method: 'modules' };
                jsonResponse['releases'] = [];

                $.each(importGrid.dataSource.view(), (index,value) => {
                    let dataItem = appendGrid.dataSource.view().find((o:any) => parseInt(o.Title) === parseInt(value.VASI_x0020_Id));
                    if (dataItem !== undefined) {
                        jsonResponse['releases'].push({
                            name: value.System_x0020_Name,
                            organization: store.value.organization,
                            version: dataItem.codeVersion,
                            status: value.System_x0020_Status,
                            permissions: { usageType: dataItem.usageType, licenses: [{ name: (dataItem.licenseName !== null) ? dataItem.licenseName : '', opRL: (dataItem.opRL !== null) ? dataItem.opRL : '' }]},
                            homepageURL: (dataItem.homepageURL !== null) ? dataItem.homepageURL : '',
                            downloadURL: (dataItem.downloadURL !== null) ? dataItem.downloadURL : '',
                            disclaimerURL: (dataItem.disclaimerURL !== null) ? dataItem.disclaimerURL : '',
                            repositoryURL: (dataItem.repositoryURL !== null) ? dataItem.repositoryURL : '',
                            vcs: dataItem.vcs,
                            laborHours: dataItem.laborHours,
                            tags: [ (dataItem.tags !== null) ? dataItem.tags : '' ],
                            languages: (value.Technology_x0020_Components !== null) ? value.Technology_x0020_Components : '',
                            contact: { name: store.value.contactName, email: store.value.contactEmail },
                            date: { created: dataItem.Created, lastModified: dataItem.Modified, metadataLastUpdated: dataItem.Modified }
                        });
                    } else {
                        jsonResponse['releases'].push({
                            name: value.System_x0020_Name,
                            organization: store.value.organization,
                            version: '',
                            status: value.System_x0020_Status,
                            permissions: { usageType: 'governmentWideReuse', licenses: [{ name: '', opRL: '' }]},
                            homepageURL: '',
                            downloadURL: '',
                            disclaimerURL: '',
                            repositoryURL: '',
                            vcs: '',
                            laborHours: 0,
                            tags: [ '' ],
                            languages: (value.Technology_x0020_Components !== null) ? value.Technology_x0020_Components : '',
                            contact: { name: store.value.contactName, email: store.value.contactEmail },
                            date: { created: value.Created, lastModified: value.Modified, metadataLastUpdated: value.Modified }
                        });
                    }
                });
                
                var a = document.createElement('a');
                var file = new Blob([JSON.stringify(jsonResponse)], {type: 'application/json'});
                a.href = URL.createObjectURL(file);
                a.download = 'code.JSON';
                a.click();
            });

        });

        return SPA.instance;
    }
}
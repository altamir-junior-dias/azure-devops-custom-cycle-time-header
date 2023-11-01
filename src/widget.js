(() => {
    let $title = $('#title');
    let $message = $('#message');
    let $counter = $('#counter');

    const getCycleTime = (items, startField, finishField) => {
        let filteredItems = items
            .filter(item => {
                return item[startField] !== undefined && item[finishField] !== undefined;
            });

        if (filteredItems.length == 0 ) {
            return 0;
        }

        return filteredItems
            .map(item => {
                if (startField == '' || finishField == '') {
                    return 0;
                }

                let startValue = new Date(item[startField]).getTime();
                let finishValue = new Date(item[finishField]).getTime();                
                let diff = Math.abs(finishValue - startValue);

                return Math.ceil(diff / (1000 * 3600 * 24));                 
            })
            .filter(value => value != null)
            .reduce((partial, value) => partial + value, 0) / filteredItems.length;
    };

    const getData = (settings) => {
        let deferred = $.Deferred();
        
        if (settings.source == 'backlog') {
            getDataFromBacklog(settings).then(data => deferred.resolve(data));
        } else {
            getDataFromQuery(settings).then(data => deferred.resolve(data));
        }

        return deferred.promise();
    };

    const getDataFromBacklog = (settings) => {
        let deferred = $.Deferred();

        AzureDevOps.Backlogs.getWorkItemTypes(settings.team, settings.backlog).then(workItemTypes => {
            let workItemTypesFilter = workItemTypes.map(workItemType => `'${workItemType}'`).join(',');

            let query = {
                wiql: `SELECT [${settings.startField}], [${settings.finishField}] ` +
                    `FROM WorkItems ` +
                    `WHERE [System.WorkItemType] in (${workItemTypesFilter})` +
                    `  AND [${settings.startField}] <> '' ` +
                    `  AND [${settings.finishField}] <> '' ` +
                    `  AND [${settings.finishField}] >= @today - ${settings.daysBehind}`,
                type: 1
            };

            AzureDevOps.Queries.getItems(query).then(items => {
                deferred.resolve(getCycleTime(items, settings.startField, settings.finishField));
            });
        });

        return deferred.promise();
    };

    const getDataFromQuery = (settings) => {
        let deferred = $.Deferred();

        AzureDevOps.Queries.getById(settings.query).then(query => {
            AzureDevOps.Queries.getItems(query).then(items => {
                deferred.resolve(getCycleTime(items, settings.startField, settings.finishField));
            });
        });

        return deferred.promise();
    };

    const getSettings = (widgetSettings) => {
        var settings = JSON.parse(widgetSettings.customSettings.data);

        return {
            title: settings?.title ?? 'Custom Cycle Time Header',
            message: settings?.message ?? '',
            source: settings?.source ?? 'backlog',
            query: settings?.query ?? '770583d1-6cec-44ad-841f-823ed722ddb1',
            startField: settings?.startField ?? '',
            finishField: settings?.finishField ?? '',
            team: settings?.team ?? VSS.getWebContext().team.id,
            backlog: settings?.backlog ?? 'Microsoft.TaskCategory',
            daysBehind: settings?.daysBehind ?? 30
        };
    };

    const load = (widgetSettings) => {
        var settings = getSettings(widgetSettings);

        $title.text(settings.title);
        $message.text(settings.message);

        getData(settings).then(cycleTime => {
            $counter.text(cycleTime.toFixed(settings.DecimalPlaces));
        });
    };

    window.LoadWidget = load;
})();
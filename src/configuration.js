(() => {
    let $context = null;
    let $widgetEvent = null;

    let $title = $('#title');
    let $message = $('#message');
    let $source = $('#source');
    let $team = $('#team');
    let $backlog = $('#backlog');
    let $query = $('#query');
    let $daysBehind = $('#days-behind');
    let $startField = $('#start-field');
    let $finishField = $('#finish-field');
    let $decimalPlaces = $('#decimal-places');

    let $teamArea = $('#team-area');
    let $backlogArea = $('#backlog-area');
    let $queryArea = $('#query-area');
    let $daysBehindArea = $('#days-behind');

    const addQueryToSelect = (query, level) => {
        level = level ?? 0;

        if (query.isFolder ?? false) {
            $query.append($('<option>')
                .val(query.id)
                .html('&nbsp;&nbsp;'.repeat(level) + query.name)
                .attr('data-level', '0')
                .css('font-weight', 'bold')
                .attr('disabled', 'disabled'));

            if (query.children.length > 0)
            {
                query.children.forEach(innerQuery => {
                    addQueryToSelect(innerQuery, level + 1);
                });
            }

        } else {
            $query.append($('<option>')
                .val(query.id)
                .html('&nbsp;&nbsp;'.repeat(level) + query.name)
                .attr('data-level', level));
        }
    };

    const changeBacklog = (notifyWidget) => {
        let deferred = $.Deferred();

        notifyWidget = notifyWidget ?? false;

        AzureDevOps.Backlogs.getFields($team.val(), $backlog.val()).then(fields => {
            updateFields(fields);

            if (notifyWidget) {
                changeSettings();
            }

            deferred.resolve();
        });

        return deferred.promise();
    };

    const changeQuery = (notifyWidget) => {
        let deferred = $.Deferred();

        notifyWidget = notifyWidget ?? false;

        AzureDevOps.Queries.getFields($query.val()).then(fields => {
            updateFields(fields);

            if (notifyWidget) {
                changeSettings();
            }

            deferred.resolve();
        });        
        
        return deferred.promise();
    };

    const changeSettings = () => {
        let eventName = $widgetEvent.ConfigurationChange;
        let eventArgs = $widgetEvent.Args(getSettingsToSave());
        $context.notify(eventName, eventArgs);
    };

    const changeSource = (notifyWidget) => {
        notifyWidget = notifyWidget ?? false;

        if ($source.val() == 'backlog') {
            $teamArea.show();
            $backlogArea.show();
            $queryArea.hide();
            $daysBehindArea.show();
        } else {
            $teamArea.hide();
            $backlogArea.hide();
            $queryArea.show();
            $daysBehindArea.hide();
        }

        if (notifyWidget) {
            changeSettings();
        }
    };

    const changeTeam = (notifyWidget) => {
        let deferred = $.Deferred();

        notifyWidget = notifyWidget ?? false;

        AzureDevOps.Backlogs.getAll($team.val()).then(backlogs => {
            $backlog.html('');
            backlogs.forEach(backlog => $backlog.append($('<option>').val(backlog.id).html(backlog.name)));

            if (notifyWidget) {
                changeSettings();
            }

            deferred.resolve();
        });

        return deferred.promise();
    };

    const loadConfiguration = (settings, context, widgetEvent) => {
        $context = context;
        $widgetEvent = widgetEvent;

        prepareControls(getSettings(settings));
    };

    const getSettings = (widgetSettings) => {
        let settings = JSON.parse(widgetSettings.customSettings.data);

        return {
            title: settings?.title ?? 'Custom Cycle Time Header',
            message: settings?.message ?? '',
            source: settings?.source ?? 'backlog',
            team: settings?.team ?? VSS.getWebContext().team.id,
            backlog: settings?.backlog ?? 'Microsoft.TaskCategory',
            query: settings?.query ?? '770583d1-6cec-44ad-841f-823ed722ddb1',
            startField: settings?.startField ?? '',
            finishField: settings?.finishField ?? '',
            daysBehind: settings?.daysBehind ?? 30,
            decimalPlaces: settings?.decimalPlaces ?? 0
        };
    };

    const getSettingsToSave = () => {
        return {
            data: JSON.stringify({
                title: $title.val(),
                message: $message.val(),
                source: $source.val(),
                team: $team.val(),
                backlog: $backlog.val(),
                query: $query.val(),
                startField: $startField.val(),
                finishField: $finishField.val(),
                daysBehind: $daysBehind.val(),
                decimalPlaces: $decimalPlaces.val()
            })
        };
    };

    const prepareControls = (settings) => {
        $title.on('change', changeSettings);
        $message.on('change', changeSettings);
        $source.on('change', changeSource);
        $team.on('change', changeTeam);
        $backlog.on('change', changeBacklog);
        $query.on('change', changeQuery);
        $startField.on('change', changeSettings);
        $finishField.on('change', changeSettings);
        $daysBehind.on('change', changeSettings);
        $decimalPlaces.on('change', changeSettings);

        let deferreds = [];
        deferreds.push(AzureDevOps.Teams.getAll());
        deferreds.push(AzureDevOps.Queries.getAllShared());

        Promise.all(deferreds).then(results => {
            let teams = results[0];
            teams.forEach(team => $team.append($('<option>').val(team.id).html(team.name)));

            let queries = results[1];
            $query.append($('<option>'));
            queries.forEach(query => addQueryToSelect(query));

            setValues(settings);
        });
    };

    const setValues = (settings) => {
        $title.val(settings.title);
        $message.val(settings.message);

        $source.val(settings.source);
        changeSource(false);

        if (settings.source == 'backlog') {
            $team.val(settings.team);

            changeTeam(false).then(_ => {
                $backlog.val(settings.backlog);
                changeBacklog(false).then(_ => {

                    $startField.val(settings.startField);
                    $finishField.val(settings.finishField);

                    $daysBehind.val(settings.daysBehind);
                });
            });
    
        } else {
            $query.val(settings.query);

            changeQuery(false).then(_ => {
                $startField.val(settings.startField);
                $finishField.val(settings.finishField);
            });    
        }

        $decimalPlaces.val(settings.decimalPlaces);
    };

    const updateFields = (fields) => {
        $startField.html('');
        $finishField.html('');

        fields.sort((a, b) => a.name > b.name ? 1 : a.name < b.name ? -1 : 0);

        fields
            .filter(field => field.type == 2)
            .forEach(field => {
                $startField.append($('<option>').val(field.referenceName).html(field.name));

                $finishField.append($('<option>').val(field.referenceName).html(field.name));
            });
    };

    window.LoadConfiguration = loadConfiguration;
    window.GetSettingsToSave = getSettingsToSave;
})();

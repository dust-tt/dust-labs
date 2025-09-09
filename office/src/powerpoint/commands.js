Office.onReady(() => {
    // Office is ready
});

function showTaskpane(event) {
    Office.ribbon.requestUpdate({
        tabs: [
            {
                id: "TabHome",
                controls: [
                    {
                        id: "TaskpaneButton",
                        enabled: true
                    }
                ]
            }
        ]
    });
    
    event.completed();
}

Office.actions.associate("ShowTaskpane", showTaskpane);
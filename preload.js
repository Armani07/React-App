
const { ipcRenderer: ipc, remote } = require('electron');

window.onload = function () {
    var tScript = document.createElement("script");
    var script = document.createElement("script");
    script.src = "js/jquery-3.2.1.min.js";
    tScript.src = "js/client.js";
    tScript.onload = script.onreadystatechange = function () {
        $(document).ready(function () {
            window.Trello.authorize({
                type: 'popup',
                name: 'Visiual Database Assurer',
                scope: {
                    read: 'true',
                    write: 'true'
                },
                expiration: '1hour',
                success: authenticationSuccess,
                error: authenticationFailure,
                persist: false
            });
            function authenticationSuccess() {
                tryIt(Trello.token());
            }

            function authenticationFailure() {
            }
        });
    };
    document.body.appendChild(script);
    document.body.appendChild(tScript);
};

function tryIt(token) {
    ipc.send('gotToken', token)
}
import { appId, redirectUri } from '../config/constants';
import Rx from 'rxjs/Rx';
import * as Msal from 'msal';

const title = document.querySelector('.title');
const statebtn = document.querySelector('.statebtn');
const mailbtn = document.querySelector('.mailbtn');
const mailcontainer = document.querySelector('.mailcontainer');
const mailtable = document.querySelector('.mailtable');

const mailBtnClickStream = Rx.Observable.fromEvent(mailbtn, 'click');
const stateBtnClickStream = Rx.Observable.fromEvent(statebtn, 'click');
const windowHashChangeStream = Rx.Observable.fromEvent(window, 'hashchange');

const graphAPIMailEndpoint = "https://graph.microsoft.com/v1.0/me/messages";
const graphAPIScopes = ["https://graph.microsoft.com/user.read", "https://graph.microsoft.com/mail.read"];
const userAgentApplication = new Msal.UserAgentApplication(appId, null, null, {
    redirectUri: redirectUri,
    cacheLocation: 'localStorage'
});

//stream that handles user input for sign in/sign out
const authorizedStream = stateBtnClickStream.flatMap(() => { 
    if (statebtn.textContent == "Sign in"){
        if (!userAgentApplication.isCallback(window.location.hash) && window.parent === window && !window.opener) {
            if (!userAgentApplication.getUser()){
                statebtn.textContent = "Signing you in";
                return Rx.Observable.of(userAgentApplication.loginRedirect(graphAPIScopes));
            } else {
                statebtn.textContent = "Sign out";
                return Rx.Observable.fromPromise(userAgentApplication.acquireTokenSilent(graphAPIScopes));
            }
        }
    } else if (statebtn.textContent == "Sign out"){
        statebtn.textContent = "Sign in";
        mailcontainer.style.visibility = "hidden";
        return Rx.Observable.of(userAgentApplication.logout());
    }
});

//stream that handles page load
const newPageStream = Rx.Observable.of({}).flatMap(() => {
    if (userAgentApplication.getUser()){
        statebtn.textContent = "Sign out";
        return Rx.Observable.fromPromise(userAgentApplication.acquireTokenSilent(graphAPIScopes));
    } else {
        mailcontainer.style.visibility = "hidden"; //on initial load
        return Rx.Observable.never();
    }
});

//stream that watches for url changes
const reloadStateStream = windowHashChangeStream.flatMap((change) => {
    if (event.oldURL.includes("#") && event.oldURL.split("#")[1].length > 0){
        //force reload
        window.location.reload(true);
    }
    return Rx.Observable.never(); 
});

//stream of auth tokens 
const userAuthStream = Rx.Observable.merge(authorizedStream, reloadStateStream, newPageStream)
                    .filter((response) => response != null);

const tokenStream = userAuthStream.map((response, err) => {
    if (!err){
        let token = response;
        mailcontainer.style.visibility = "visibile";

        return Rx.Observable.of(token);
    } else {
        return Rx.Observable.never();
    }
});

//stream of respones from Graph API
const responseStream = tokenStream.flatMap((token)=>{
    const headers = {"Authorization": "Bearer "+token.value};
    return Rx.Observable.fromPromise(fetch(graphAPIMailEndpoint, {headers: headers}))
}).combineLatest(mailBtnClickStream, 
    (resp, click) => { return resp });

//stream of deserialized responses
const mailJsonStream = responseStream.flatMap((resp)=>{
    return Rx.Observable.fromPromise(resp.json());
})

mailJsonStream.subscribe((json)=>{
    for (let[index, message] of json.value.entries()){
        let row = mailtable.insertRow(index+1);
        let senderCell = row.insertCell(0);
        let subjectCell = row.insertCell(1);
        let previewCell = row.insertCell(2);
        let isReadCell = row.insertCell(3);
        senderCell.innerHTML = message.sender.emailAddress.name;
        subjectCell.innerHTML = message.subject;
        previewCell.innerHTML = message.bodyPreview;
        isReadCell.innerHTML = (message.isRead) ? "Read" : "Not read";
    }
});
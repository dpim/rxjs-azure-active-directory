
import { appId, redirectUri } from '../config/constants';
import Rx from 'rxjs/Rx';
import * as Msal from 'msal';

const graphAPIScopes = ["https://graph.microsoft.com/user.read", "https://graph.microsoft.com/mail.read"];
const userAgentApplication = new Msal.UserAgentApplication(appId, null, null, {
    redirectUri: redirectUri,
    cacheLocation: 'localStorage'
});

const signInButton = document.querySelector('#signInButton');
const signOutButton = document.querySelector('#signOutButton');

const signInButtonClickStream = Rx.Observable.fromEvent(signInButton, 'click');
const signOutButtonClickStream = Rx.Observable.fromEvent(signOutButton, 'click');
const windowHashChangeStream = Rx.Observable.fromEvent(window, 'hashchange');

const userStream = Rx.Observable.of(userAgentApplication.getUser());

const authorizedStream = userStream.flatMap((user)=>{
    showSignedOutState();
    if (user){
        return Rx.Observable.fromPromise(userAgentApplication.acquireTokenSilent(graphAPIScopes));
    } else {
        return Rx.Observable.never();
    }
}).filter(response => response != null);

export const tokenStream = authorizedStream.flatMap((response, err) => {
    if (!err){
        showSignedInState();
        return Rx.Observable.of(response);
    } else {
        return Rx.Observable.never();
    }
});

//handle sign in/sign out events
signOutButtonClickStream.subscribe(() => {
    if (userAgentApplication.getUser()){
        userAgentApplication.logout();
    }
});

signInButtonClickStream.subscribe(() => {
    userAgentApplication.loginRedirect(graphAPIScopes);
});

//watch for redirect
windowHashChangeStream.subscribe((change) => {
    if (event.oldURL.includes("#") && event.oldURL.split("#")[1].length > 0){
        //force reload
        window.location.reload(true);
    }
});

//aux functions
function showSignedInState(){
    signInButton.classList.add("hidden");
    signOutButton.classList.remove("hidden");
}

function showSignedOutState(){
    signInButton.classList.remove("hidden");
    signOutButton.classList.add("hidden");
}

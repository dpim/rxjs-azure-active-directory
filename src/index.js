
import Rx from 'rxjs/Rx';
import { tokenStream } from './auth.js';

const mailButton = document.querySelector('#mailButton');
const mailContainer = document.querySelector('#mailContainer');
const mailTable = document.querySelector('#mailTable');

const mailButtonClickStream = Rx.Observable.fromEvent(mailButton, 'click');
const graphAPIMailEndpoint = "https://graph.microsoft.com/v1.0/me/messages";

//stream of responses from Graph API
const responseStream = tokenStream.flatMap((token) => {
    const headers = {"Authorization": "Bearer " + token};
    return Rx.Observable.fromPromise(
        fetch(graphAPIMailEndpoint, { headers: headers })
    );
}).combineLatest(mailButtonClickStream, 
    (resp, click) => { if (click) return resp });

//stream of deserialized responses
const mailJsonStream = responseStream.flatMap((resp) => {
    return (resp.bodyUsed) ? Rx.Observable.never() : Rx.Observable.fromPromise(resp.json());
})

mailJsonStream.subscribe((json) => {
    for (let[index, message] of json.value.entries()){
        let row = mailTable.insertRow(index+1);
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

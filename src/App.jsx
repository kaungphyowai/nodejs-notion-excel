import React, { useState } from "react";
import { AuthenticatedTemplate, UnauthenticatedTemplate, useMsal } from "@azure/msal-react";
import { loginRequest } from "./authConfig";
import { PageLayout } from "./components/PageLayout";
import { ProfileData } from "./components/ProfileData";
import { callMsGraph } from "./graph";
import Button from "react-bootstrap/Button";
import "./styles/App.css";
const { Client } = require("@notionhq/client")

// Initializing a client
const notion = new Client({
    auth: process.env.REACT_APP_NOTION_KEY,
})


/**
 * Renders information about the signed-in user or a button to retrieve data about the user
 */
const ProfileContent = () => {
    const { instance, accounts } = useMsal();
    const [graphData, setGraphData] = useState(null);

    function RequestProfileData() {
        // Silently acquires an access token which is then attached to a request for MS Graph data
        instance.acquireTokenSilent({
            ...loginRequest,
            account: accounts[0]
        }).then((response) => {
            callMsGraph(response.accessToken).then(response => setGraphData(response));
        });
    }

    /**
     * 
     * @returns {Promise<Response>} Promise Object
     */
    function getNotionData() {
        const option = {
            mode: 'no-cors'
        }
        return fetch('http://localhost:5000/getNotionDatabase').then(data => data.json()).then(final => final)
    }

    function Sync() {
        getNotionData().then(data => {
            //start the conversation
            //find out about columns name and its type
            let allData = data.results;
            let demoData = allData[0].properties;
            let allColumns = Object.keys(demoData);
            let titleColumn = allColumns.filter(column => demoData[column].type === 'title')[0];
            let remainingColumn = allColumns.filter(column => demoData[column].type !== 'title');
            let remianingColandType = remainingColumn.map(column => {
                return {
                    columnName: column,
                    type: demoData[column].type
                }
            })
            //make an conversation algorithm
            //prepare header
            let finalCSV = titleColumn;
            for (let col in remainingColumn) {
                finalCSV = finalCSV + "," + remainingColumn[col];
            }
            finalCSV = finalCSV + '\n';
            console.log(allData.lenght)
            //make the body
            for (let i = 0; i < allData.length; i++) {
                let data = allData[i].properties;
                //add the title data first
                finalCSV = finalCSV + data[titleColumn].title[0].text.content;
                //add the remaing data
                for (let j = 0; j < remianingColandType.length; j++) {
                    let { columnName, type } = remianingColandType[j]
                    if (remianingColandType[j].type === 'rich_text') {
                        finalCSV = finalCSV + ',' + data[columnName][type][0].text.content;
                    }
                    if (remianingColandType[j].type === 'multi_select') {
                        finalCSV = finalCSV + ',';
                        for (let k = 0; k < data[columnName][type].length; k++) {
                            if(k == 0 ){

                                finalCSV = finalCSV  + data[columnName][type][k].name; 
                                break;
                            }
                            finalCSV = finalCSV + " " + data[columnName][type][k].name;
                            
                        }
                    }
                }
                finalCSV = finalCSV + '\n';
            }

            //update the file in office
            instance.acquireTokenSilent({
                ...loginRequest,
                account: accounts[0]
            }).then(response => {
                const headers = new Headers();
                const bearer = `Bearer ${response.accessToken}`;
                headers.append("Authorization", bearer);
                headers.append('Content-Type', "text/plain");

                const options = {
                    method: "PUT",
                    headers: headers,
                    body: finalCSV
                };
                fetch("https://graph.microsoft.com/v1.0/me/drive/items/01WOPEE4MRC4BINADKBRB2PBP7TMYMD5WR/content", options)
                    .then(response => console.log("Name Changes Complete" + response))
                    .catch(error => console.log(error));
            })


        })
    }
    return (
        <>
            <h5 className="card-title">Welcome {accounts[0].name}</h5>
            {graphData ?
                <ProfileData graphData={graphData} />
                :
                <Button variant="secondary" onClick={RequestProfileData}>Request Profile Information</Button>
            }
            <Button onClick={Sync} >Sync For the Bloody Hell</Button>
        </>
    );
};

/**
 * If a user is authenticated the ProfileContent component above is rendered. Otherwise a message indicating a user is not authenticated is rendered.
 */
const MainContent = () => {
    return (
        <div className="App">
            <AuthenticatedTemplate>
                <ProfileContent />
            </AuthenticatedTemplate>

            <UnauthenticatedTemplate>
                <h5 className="card-title">Please sign-in to see your profile information.</h5>
            </UnauthenticatedTemplate>
        </div>
    );
};

export default function App() {
    return (
        <PageLayout>
            <MainContent />
        </PageLayout>
    );
}

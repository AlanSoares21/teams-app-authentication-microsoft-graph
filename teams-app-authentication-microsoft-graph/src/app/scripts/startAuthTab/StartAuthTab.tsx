import * as React from "react";
import { Provider, Text, Button } from "@fluentui/react-northstar";
import TeamsBaseComponent, { ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import {stringify} from 'querystring';
/**
 * State for the startAuthTabTab React component
 */
export interface IStartAuthTabState extends ITeamsBaseComponentState {
    entityId?: string;
}

/**
 * Properties for the startAuthTabTab React component
 */
export interface IStartAuthTabProps {

}

/**
 * Implementation of the startAuth content page
 */
export class StartAuthTab extends TeamsBaseComponent<IStartAuthTabProps, IStartAuthTabState> {

    public async componentWillMount() {
        microsoftTeams.initialize();
    }

    public authOpenWindow = async()=>{
        microsoftTeams.initialize()
        microsoftTeams.getContext((context)=>{
            // Go to the Azure AD authorization endpoint
            let queryParams = {
                client_id: process.env.AUTH_APP_ID,
                response_type: "code",
                response_mode: "fragment",
                scope: "offline_access user.read mail.read",
                redirect_uri: window.location.origin + "/endAuthTab",
            };
        
            let authorizeEndpoint = "https://login.microsoftonline.com/" + (context.tid? context.tid : 'commom') + "/oauth2/v2.0/authorize?" + stringify(queryParams);
            window.location.assign(authorizeEndpoint);
        });
    }

    /**
     * The render() method to create the UI of the tab
     */
    public render() {
        return (
            <Provider >
                <Button content="Start" onClick={this.authOpenWindow}/>
                <br></br>
                <Text>
                    Click
                </Text>
            </Provider>
        );
    }
}

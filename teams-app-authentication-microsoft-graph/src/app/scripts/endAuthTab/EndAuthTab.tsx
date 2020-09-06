import * as React from "react";
import { Provider, Text } from "@fluentui/react-northstar";
import TeamsBaseComponent, { ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
/**
 * State for the endAuthTabTab React component
 */
export interface IEndAuthTabState extends ITeamsBaseComponentState {
    entityId?: string;
}

/**
 * Properties for the endAuthTabTab React component
 */
export interface IEndAuthTabProps {

}

/**
 * Implementation of the endAuth content page
 */
export class EndAuthTab extends TeamsBaseComponent<IEndAuthTabProps, IEndAuthTabState> {

    public async componentWillMount() {
        microsoftTeams.initialize();
        
        const hashCodeParts = window.location.hash.split('&')[0].slice(1).split('=');
        
        if (hashCodeParts[0] === 'code' && hashCodeParts.length === 2) {
            // Authentication/authorization sucess
            microsoftTeams.authentication.notifySuccess(hashCodeParts[1]);
        }
        else {
            // Unexpected condition: hash does not contain error or access_token parameter
            microsoftTeams.authentication.notifyFailure(hashCodeParts.join('='));
        }
        microsoftTeams.navigateToTab({tabName:'auth',entityId:'b0724e3f-6c12-4440-9617-162ea40d1309'})
    }

    /**
     * The render() method to create the UI of the tab
     */
    public render() {
        return (
            <Provider >
                <Text>
                    ...
                </Text>
                <br></br>
            </Provider>
        );
    }
}

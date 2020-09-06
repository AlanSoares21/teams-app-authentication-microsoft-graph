import * as React from "react";
import { Provider, Flex, Text, Button, Header } from "@fluentui/react-northstar";
import TeamsBaseComponent, { ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";

import { getToken } from '../../token-request';

/**
 * State for the authTabTab React component
 */
export interface IAuthTabState extends ITeamsBaseComponentState {
    tenant_id?: string;
    response: string;
}

/**
 * Properties for the authTabTab React component
 */
export interface IAuthTabProps {

}

/**
 * Implementation of the auth content page
 */
export class AuthTab extends TeamsBaseComponent<IAuthTabProps, IAuthTabState> {
    
    public async componentWillMount() {
        this.setState({
            response: 'faça o request'
        });
        
        if (await this.inTeams()) {
            microsoftTeams.initialize();
            microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);
            microsoftTeams.getContext((context) => {
                microsoftTeams.appInitialization.notifySuccess();
                this.setState({
                    tenant_id: context.tid
                });
                this.updateTheme(context.theme);
            });
        }
    }

    private authSucess = async (code: string)=>{
        alert('sucess');
        const redirect_uri = window.location.origin + '/endAuthTab'
        if(this.state.tenant_id){
            const tokenResponse = await getToken( code, redirect_uri, this.state.tenant_id);
            this.setState({
                response: JSON.stringify(tokenResponse,null,'\n')
            })
        }
        
    }

    private authFail = (reason: string)=>{
        alert(reason);
    }

    private authStart = async () =>{
        microsoftTeams.authentication.authenticate({
            url: window.location.origin + "/startAuthTab",
            successCallback:  this.authSucess,
            failureCallback: this.authFail
        })           
    }
    
    
    public render() {
        return (
            <Provider >
                <Text>
                    {this.state.response}
                </Text>
                <br></br>
                <Button content="init auth" onClick={this.authStart} />
            </Provider>
        );
    }
}

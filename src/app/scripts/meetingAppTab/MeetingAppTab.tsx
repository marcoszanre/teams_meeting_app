import * as React from "react";
import { Provider, Flex, Text, Button, Header } from "@fluentui/react-northstar";
import TeamsBaseComponent, { ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import jwt_decode from "jwt-decode";
/**
 * State for the meetingAppTabTab React component
 */
export interface IMeetingAppTabState extends ITeamsBaseComponentState {
    entityId?: string;
    name?: string;
    error?: string;
    frameContext?: string;
    meetingId?: string;
    themeString?: string;
    tid?: string;
    userPrincipalName?: string;
    userObjectId?: string;
    role?: string;
    chatId?: string;
}

/**
 * Properties for the meetingAppTabTab React component
 */
export interface IMeetingAppTabProps {

}

/**
 * Implementation of the Meeting App Tab content page
 */
export class MeetingAppTab extends TeamsBaseComponent<IMeetingAppTabProps, IMeetingAppTabState> {

    public async componentWillMount() {
        this.updateTheme(this.getQueryVariable("theme"));


        microsoftTeams.initialize(() => {
            microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);
            microsoftTeams.getContext((context) => {
                this.setState({
                    entityId: context.entityId
                });
                this.updateTheme(context.theme);
                microsoftTeams.authentication.getAuthToken({
                    successCallback: (token: string) => {
                        const decoded: { [key: string]: any; } = jwt_decode(token) as { [key: string]: any; };
                        this.setState({ name: decoded!.name });
                        microsoftTeams.appInitialization.notifySuccess();
                    },
                    failureCallback: (message: string) => {
                        this.setState({ error: message });
                        microsoftTeams.appInitialization.notifyFailure({
                            reason: microsoftTeams.appInitialization.FailedReason.AuthFailed,
                            message
                        });
                    },
                    resources: [process.env.MEETINGAPPTAB_APP_URI as string]
                });
            });
        });

        microsoftTeams.getContext(async (ctx) => {

            console.log(ctx);

            this.setState( {
                frameContext: ctx.frameContext,
                meetingId: ctx.meetingId,
                themeString: ctx.theme,
                tid: ctx.tid,
                userPrincipalName: ctx.userPrincipalName,
                userObjectId: ctx.userObjectId,
                chatId: ctx.chatId
            });

            const res = await fetch(`/api/role?meetingId=${ctx.meetingId}&userId=${ctx.userObjectId}`);
            const json = await res.json();
            this.setState( { role: json.role } );
            });
    }

    /**
     * The render() method to create the UI of the tab
     */
    public render() {
        return (
            <Provider theme={this.state.theme}>
                <Flex fill={true} column styles={{
                    padding: ".8rem 0 .8rem .5rem"
                }}>
                    <Flex.Item>
                        <Header content="This is your tab" />
                    </Flex.Item>
                    <Flex.Item>
                        <div>
                            <div>
                                <Text content={`Hello ${this.state.name}`} />
                            </div>
                            <div>
                                <Text content={`frameContext: ${this.state.frameContext}`} />
                            </div>
                            <div>
                                <Text content={`meetingId: ${this.state.meetingId}`} />
                            </div>
                            <div>
                                <Text content={`theme: ${this.state.themeString}`} />
                            </div>
                            <div>
                                <Text content={`tenantId: ${this.state.tid}`} />
                            </div>
                            <div>
                                <Text content={`User UPN: ${this.state.userPrincipalName}`} />
                            </div>
                            <div>
                                <Text content={`User AAD ID: ${this.state.userObjectId}`} />
                            </div>
                            <div>
                                <Text content={`Role: ${this.state.role}`} />
                            </div>
                            <div>
                                <Text content={`ChatId: ${this.state.chatId}`} />
                            </div>
                            {this.state.error && <div><Text content={`An SSO error occurred ${this.state.error}`} /></div>}

                            <div>
                                <Button onClick={async () => await fetch(`/api/bubble?chatId=${this.state.chatId}`)}>Send Bubble</Button>
                            </div>
                        </div>
                    </Flex.Item>
                    <Flex.Item styles={{
                        padding: ".8rem 0 .8rem .5rem"
                    }}>
                        <Text size="smaller" content="(C) Copyright Contoso" />
                    </Flex.Item>
                </Flex>
            </Provider>
        );
    }
}

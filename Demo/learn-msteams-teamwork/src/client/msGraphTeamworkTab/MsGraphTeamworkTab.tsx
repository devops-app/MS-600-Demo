import * as React from "react";
import { Provider, Flex, Text, Button, Header, Avatar, List } from "@fluentui/react-northstar";
import { useState, useEffect, useCallback } from "react";
import { useTeams } from "msteams-react-base-component";
import { app, authentication } from "@microsoft/teams-js";
import jwtDecode from "jwt-decode";

/**
 * Implementation of the MSGraph Teamwork content page
 */
export const MsGraphTeamworkTab = () => {

    const [{ inTeams, theme, context }] = useTeams();
    const [entityId, setEntityId] = useState<string | undefined>();
    const [name, setName] = useState<string>();
    const [error, setError] = useState<string>();
    const [ssoToken, setSsoToken] = useState<string>();
    const [msGraphOboToken, setMsGraphOboToken] = useState<string>();
    const [photo, setPhoto] = useState<string>();
    const [joinedTeams, setJoinedTeams] = useState<any[]>();

    useEffect(() => {
        if (inTeams === true) {
            authentication.getAuthToken({
                resources: [process.env.TAB_APP_URI as string],
                silent: false
            } as authentication.AuthTokenRequestParameters).then(token => {
                const decoded: { [key: string]: any; } = jwtDecode(token) as { [key: string]: any; };
                setName(decoded!.name);
                setSsoToken(token);
                app.notifySuccess();
            }).catch(message => {
                setError(message);
                app.notifyFailure({
                    reason: app.FailedReason.AuthFailed,
                    message
                });
            });
        } else {
            setEntityId("Not in Microsoft Teams");
        }
    }, [inTeams]);

    useEffect(() => {
        if (context) {
            setEntityId(context.page.id);
        }
    }, [context]);

    const exchangeSsoTokenForOboToken = useCallback(async () => {
        const response = await fetch(`/exchangeSsoTokenForOboToken/?ssoToken=${ssoToken}`);
        const responsePayload = await response.json();
        if (response.ok) {
            setMsGraphOboToken(responsePayload.access_token);
        } else {
            if (responsePayload!.error === "consent_required") {
                setError("consent_required");
            } else {
                setError("unknown SSO error");
            }
        }
    }, [ssoToken]);

    useEffect(() => {
        // if the SSO token is defined...
        if (ssoToken && ssoToken.length > 0) {
            exchangeSsoTokenForOboToken();
        }
    }, [exchangeSsoTokenForOboToken, ssoToken]);

    const getProfilePhoto = useCallback(async () => {
        if (!msGraphOboToken) { return; }
        const endpoint = "https://graph.microsoft.com/v1.0/me/photo/$value";
        const requestObject = {
            method: "GET",
            headers: {
                accept: "image/jpg",
                authorization: `bearer ${msGraphOboToken}`
            }
        };
        const response = await fetch(endpoint, requestObject);
        if (response.ok) {
            setPhoto(URL.createObjectURL(await response.blob()));
        }
    }, [msGraphOboToken]);    

    const getJoinedTeams = useCallback(async () => {
        if (!msGraphOboToken) { return; }
        const endpoint = "https://graph.microsoft.com/v1.0/me/joinedTeams";
        const requestObject = {
            method: "GET",
            headers: {
                accept: "application/json",
                authorization: `bearer ${msGraphOboToken}`
            }
        };
        const response = await fetch(endpoint, requestObject);
        const responsePayload = await response.json();
        if (response.ok) {
            const listFriendlyJoinedTeams = responsePayload.value.map((team: any) => ({
                key: team.id,
                header: team.displayName,
                content: `Team ID: ${team.id}`
            }));
            setJoinedTeams(listFriendlyJoinedTeams);
        }
    }, [msGraphOboToken]);

    useEffect(() => {
        getJoinedTeams();
        getProfilePhoto();
    }, [getJoinedTeams, getProfilePhoto, msGraphOboToken]);

    /**
     * The render() method to create the UI of the tab
     */
    return (
        <Provider theme={theme}>
            <Flex fill={true} column styles={{
                padding: ".8rem 0 .8rem .5rem"
            }}>
                <Flex.Item>
                    <Header content="This is your tab" />
                </Flex.Item>
                <Flex.Item>
                    <div>
                        <div>
                            <Text content={`Hello ${name}`} />
                        </div>
                        {photo && <div><Avatar image={photo} size='largest' /></div>}
                        {joinedTeams && <div><h3>You belong to the following teams:</h3><List items={joinedTeams} /></div>}
                        {error && <div><Text content={`An SSO error occurred ${error}`} /></div>}

                        <div>
                            <Button onClick={() => alert("It worked!")}>A sample button</Button>
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
};

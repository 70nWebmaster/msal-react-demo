import { ProfileData } from "../components/ProfileData";
import { InteractionType } from "@azure/msal-browser";
import { useMsalAuthentication } from "@azure/msal-react";
import { useState, useEffect } from "react";

import { fetchData } from '../fetch'

export const Profile = () => {
    const [graphData, setGraphData] = useState(null);
    const { result, error } = useMsalAuthentication(InteractionType.Popup, {
        scopes: ["user.read"]
    });

    useEffect(() => {
        if (!!graphData) {
            return;
        }

        if (!!error) {
            console.log(error);
            return;
        }

        if (result) {
            const { accessToken } = result;
            fetchData('https://graph.microsoft.com/v1.0/me', accessToken)
                .then(response => setGraphData(response))
                .catch(error => console.log(error));
        }
    }, [graphData, error, result]);

    return (
        <>
            { graphData ? <ProfileData graphData={graphData} /> : null }
        </>
    )
}

import { ProfileData } from "../components/ProfileData";
import { InteractionType } from "@azure/msal-browser";
import { useMsalAuthentication } from "@azure/msal-react";
import { useState, useEffect } from "react";

import { fetchData } from "../fetch";
import { editableInputTypes } from "@testing-library/user-event/dist/utils";

export const Profile = () => {
  const [graphData, setGraphData] = useState(null);
  const { result, error } = useMsalAuthentication(InteractionType.Popup, {
    scopes: ["user.read"],
  });

  useEffect(() => {
    if (!!graphData) {
      return;
    }
    if (!!error) {
      console.log(error);
      return;
    }

    if (result) {
      const { accessToken } = result;
      fetchData("https://graph.microsoft.com/v1.0/me", accessToken)
        .then((response) => setGraphData(response))
        .catch((error) => console.log(error));
    }
  }, [graphData, error, result]);

  return <>{graphData ? <ProfileData graphData={graphData} /> : null}</>;
};

// start of edit
// Construct email object
const mail = {
  subject: "Microsoft Graph JavaScript Email",
  toRecipients: [
    {
      emailAddress: {
        address: "webmaster@support907.com",
      },
    },
  ],
  body: {
    content:
      "<h1>MicrosoftGraph JavaScript</h1>Check out https://github.com/70nWebmaster/msal-react-demo",
    contentType: "html",
  },
};
try {
  let response = await client.api("/me/sendMail").post({ message: mail });
  console.log(response);
} catch (error) {
  throw error;
}
// end of edit

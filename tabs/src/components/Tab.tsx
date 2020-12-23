// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import React, { useEffect, useState } from "react";
import "./App.css";
import * as microsoftTeams from "@microsoft/teams-js";
import { darkTheme, lightTheme, Theme } from "./Tab.style";

/**
 * The 'PersonalTab' component renders the main tab content
 * of your app.
 */
export default () => {
  const [context, setContext] = useState<microsoftTeams.Context>();
  const [newTheme, setNewTheme] = useState<Theme>({});

  let userName = "";

  const updateTheme = (theme: string | undefined) => {
    if (theme === "default") {
      setNewTheme(lightTheme);
    } else {
      setNewTheme(darkTheme);
    }
  };

  useEffect(() => {
    microsoftTeams.getContext((context: microsoftTeams.Context) => {
      if (context) {
        // Initial update to context, theme, and username
        userName = context.userPrincipalName ?? "";
        updateTheme(context.theme);
        setContext(context);
        microsoftTeams.registerOnThemeChangeHandler((theme) => {
          // Update theme when it is changed by user
          if (theme !== context.theme) {
            setContext({ ...context });
          }
          updateTheme(context.theme);
        });
      }
    });
  }, [JSON.stringify(context), JSON.stringify(newTheme)]);

  return (
    <div style={newTheme}>
      <p>Username: {userName}</p>
      <h1>Important Contacts</h1>
      <ul>
        <li>
          Help Desk:{" "}
          <a href="mailto:support@company.com">support@company.com</a>
        </li>
        <li>
          Human Resources: <a href="mailto:hr@company.com">hr@company.com</a>
        </li>
        <li>
          Facilities:{" "}
          <a href="mailto:facilities@company.com">facilities@company.com</a>
        </li>
      </ul>
    </div>
  );
};

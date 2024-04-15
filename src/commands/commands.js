Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

async function checkTokenExpirationAndRefresh() {
  // console.log("==== " + localStorage.getItem("tokens_data"));
  const tokensData = JSON.parse(localStorage.getItem("tokens_data"));
  // console.log(tokensData);

  if (!tokensData || Date.now() >= tokensData.expiresIn) {
    return await refreshToken(tokensData);
  } else {
    // Token still valid
    return tokensData;
  }
}

async function refreshToken(tokensData) {
  if (!tokensData || !tokensData.refreshToken) {
    return; // Cannot refresh without a refresh token
  }

  const myHeaders = new Headers();
  myHeaders.append("Content-Type", "application/json");
  const raw = JSON.stringify({
    refreshToken: tokensData.refreshToken,
  });

  const requestOptions = {
    method: "POST",
    headers: myHeaders,
    body: raw,
    redirect: "follow",
  };

  try {
    const response = await fetch("https://ms-tokens.vercel.app/refreshTokens", requestOptions);
    if (!response.ok) {
      throw new Error(`Failed to refresh token: ${response.statusText}`);
    }
    const result = await response.json();
    const { access_token, refresh_token, expires_in } = result.result;
    const tokens = {
      accessToken: access_token,
      refreshToken: refresh_token,
      expiresIn: Date.now() + expires_in * 1000,
    };
    localStorage.setItem("tokens_data", JSON.stringify(tokens));
    return tokens;
  } catch (error) {
    console.error(error);
  }
}

async function action(event) {
  let tokensData = await checkTokenExpirationAndRefresh();
  let token = tokensData;
  if (!token) {
    openSignInDialog(event);
  } else {
    addCategory(event, token);
  }
}

let dialog;
function openSignInDialog(event) {
  // const redirectURI = "https://localhost:3000";
  const redirectURI = "https://outlook-addin-category.vercel.app";
  const tenant = "697d166b-cfa6-4cbe-b257-7b82e6ac01f0";
  const clientId = "90df907f-ce8d-426c-977e-3f0372e2453b";

  const URL = `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/authorize?response_type=code&client_id=${clientId}&redirect_uri=${redirectURI}/assets/login.html&scope=openid+profile+offline_access+User.Read+Mail.Read+Mail.ReadWrite`;
  Office.context.ui.displayDialogAsync(URL, { height: 30, width: 30 }, function (asyncResult) {
    dialog = asyncResult.value;
    dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
  });
  function processMessage(arg) {
    const token = arg.message;
    // console.log(token);
    const myHeaders = new Headers();
    myHeaders.append("Content-Type", "application/json");
    const raw = JSON.stringify({
      code: token,
    });
    const requestOptions = {
      method: "POST",
      headers: myHeaders,
      body: raw,
    };
    fetch("https://ms-tokens.vercel.app/getAccessToken", requestOptions)
      .then((response) => {
        if (!response.ok) {
          throw new Error(`Failed to get access token: ${response.statusText}`);
        }
        return response.json();
      })
      .then(async (result) => {
        const { access_token, refresh_token, expires_in } = result.result;
        const tokens = {
          accessToken: access_token,
          refreshToken: refresh_token,
          expiresIn: Date.now() + expires_in * 1000,
        };
        localStorage.setItem("tokens_data", JSON.stringify(tokens));
        dialog.close();
        if (tokens.accessToken) {
          try {
            await addCategory(event, tokens);
            showMessage(event, "Sign-in successful, working on category, please wait...");
          } catch (error) {
            console.error(error);
            showMessage(event, "Error adding category, please try again later!");
          }
        }
      })
      .catch((error) => {
        console.error(error);
        showMessage(event, "Signin Failed, please try again later!");
        dialog.close();
      });
  }
}

async function addCategory(event, token) {
  try {
    if (!token) {
      throw new Error("Token is not available.");
    }
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$select=subject,sender&top=100`,
      {
        method: "GET",
        headers: {
          Authorization: `Bearer ${token.accessToken}`,
          "Content-Type": "application/json",
        },
      }
    );
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    const data = await response.json();
    const currentItem = Office.context.mailbox.item;
    const matchingEmails = data.value.filter((graphEmail) => {
      return (
        graphEmail.subject === currentItem.subject &&
        graphEmail.sender.emailAddress.address === currentItem.sender.emailAddress
      );
    });
    for (const email of matchingEmails) {
      await insertCategory(event, email.id, token);
    }
    event.completed();
  } catch (error) {
    console.error("Error:", error);
  }
}

async function insertCategory(event, id, token) {
  try {
    const category = "Blue category";
    const url = `https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages/${id}`;
    const body = JSON.stringify({ categories: [category] });

    const response = await fetch(url, {
      method: "PATCH",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${token.accessToken}`,
      },
      body: body,
    });

    if (!response.ok) {
      throw new Error(`Error adding category: ${response.statusText}`);
    }
    showMessage(event, "Category added successfully!");
    event.completed();
  } catch (error) {
    console.error("Error:", error);
    showMessage(event, "Error adding category");
  }
}

function showMessage(event, res) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: res,
    icon: "Icon.80x80",
    persistent: true,
  };
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);
  event.completed();
}

// Register the function with Office.
Office.actions.associate("action", action);

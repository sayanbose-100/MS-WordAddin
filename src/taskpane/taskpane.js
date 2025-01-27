const { PublicClientApplication } = require("@azure/msal-browser");

let policies = [];

const fetchPolicies = async (userEmail) => {
  const policyContainer = document.getElementById("policies_container");

  const response = await fetch(`http://localhost:3001/policies/${userEmail}`,{
    method: "GET",
    headers: {
      "Content-Type": "application/json",
    },
  });

  if (!response.ok) {
    throw new Error(`Error fetching policies: ${response.statusText}`);
  }
  const data = await response.json();
  if(data){
    data.map((ele) => policies.push(ele));
  }
  console.log("Policies: ",policies);
  if(policies.length > 0) {
    const select = document.createElement("select");
    select.id = "policies";
    select.title = "policies";

    policies.forEach((policy) => {
    const option = document.createElement("option");
    option.value = `${policy.policyId}`;
    option.innerText = `${policy.policyName}`;
    select.appendChild(option);
    // console.log(policy);
    })
    policyContainer.appendChild(select);
  }
}

const msalConfig = {
  auth: {
    clientId: "d2cbb35c-ca9b-4927-809e-6deb767ab582",
    authority: "https://login.microsoftonline.com/49bdfd6a-1abb-4a96-afec-ce99fe8a15c1",
    redirectUri: "https://localhost:3000/",
  },
};

const msalInstance = new PublicClientApplication(msalConfig);

async function initializeMsal() {
  await msalInstance.initialize();
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("create").onclick = create;
    document.getElementById("sync").onclick = syncF;
    document.getElementById("emailFetch").onclick = getUserEmail;
  }
});

export async function create() {
  return Word.run(async (context) => {
    const policyOwner = document.getElementById("policyOwner");
    const policyName = document.getElementById("policyName");

    policyOwner.style.fontSize = "20px";
    policyName.style.color = "blue";

    // policy entry
    const policyEntry = `Name of the Owner: ${policyOwner.value} \\n Name of the Policy: ${policyName.value}`;

    // Insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph(`${policyEntry}`, Word.InsertLocation.end);

    await context.sync();
  });
}

export async function syncF() {
  const syncMessage = document.getElementById("sync_message");
  const policyList = document.getElementById("policy-list");
  initializeMsal()
    .then(() => {
      return Word.run(async (context) => {
        try {
          const loginResponse = await msalInstance.loginPopup({
            scopes: ["Files.ReadWrite.All"],
          });

          const account = msalInstance.getAccount(loginResponse.account.username);
          msalInstance.setActiveAccount(account);

          const tokenResponse = await msalInstance.acquireTokenSilent({
            scopes: ["Files.ReadWrite.All"],
            account: account,
          });

          const accessToken = tokenResponse.accessToken;

          // Use the access token to access OneDrive
          console.log("Access token acquired:", accessToken);
          syncMessage.innerText = "Token acquired";
          syncMessage.style.color = "green";

          // Fetch all MS Word files from the user's OneDrive
          const wordFiles = await fetchCurrentWordFiles(accessToken);
          console.log("Word files:", wordFiles);
          saveDocumentCredentials(wordFiles);
          console.log("Sent the files to the Server");
          // wordFiles.forEach((ele) => {
          //   const li = document.createElement("li");
          //   li.innerText = ele;
          //   policyList.appendChild(li);
          // });
        } catch (error) {
          console.error("Authentication error:", error);
        }
      });
    })
    .catch(() => {
      console.error("MSAL initialization failed");
    });
}

async function fetchCurrentWordFiles(accessToken) {
  const endpoint = `https://graph.microsoft.com/v1.0/me/drive/recent`;

  const response = await fetch(endpoint, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
  });

  if (!response.ok) {
    throw new Error(`Error fetching Word files: ${response.statusText}`);
  }

  const data = await response.json();
  // const wordFileLinks = data.value.map((file) => file.webUrl);
  // return wordFileLinks;
  // return {
  //   id: data.id,
  //   name: data.name,
  //   webUrl: data.webUrl,
  // }
  // return data;
  return {
    id: data.value[0].id,
    webUrl: data.value[0].webUrl,
  };
}

async function saveDocumentCredentials(data) {
  try {
    const response = await fetch("http://localhost:3001/add", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(data),
    });

    if (!response.ok) {
      throw new Error(`Error saving document credentials: ${response.statusText}`);
    }

    const responseData = await response.json();
    console.log("Successfully saved document credentials", responseData);
  } catch (err) {
    console.error(err);
  }
}

async function getUserEmail(accessToken) {
  return Word.run(async (context) => {
    initializeMsal()
      .then(() => {
        return Word.run(async (context) => {
          try {
            const loginResponse = await msalInstance.loginPopup({
              scopes: ["Files.ReadWrite.All"],
            });

            const account = msalInstance.getAccount(loginResponse.account.username);
            msalInstance.setActiveAccount(account);

            const tokenResponse = await msalInstance.acquireTokenSilent({
              scopes: ["Files.ReadWrite.All"],
              account: account,
            });

            const accessToken = tokenResponse.accessToken;

            const endpoint = `https://graph.microsoft.com/v1.0/me`;

            const response = await fetch(endpoint, {
              method: "GET",
              headers: {
                Authorization: `Bearer ${accessToken}`,
                "Content-Type": "application/json",
              },
            });
            if (!response.ok) {
              throw new Error(`Error fetching user email: ${response.statusText}`);
            }
            const data = await response.json();
            const userEmail = data.mail || data.userPrincipalName;
            console.log("User email: " + userEmail);
            fetchPolicies(userEmail);
          } catch (error) {
            console.error("Authentication error:", error);
          }
        });
      })
      .catch(() => {
        console.error("MSAL initialization failed");
      });
  });
}

const { PublicClientApplication } = require("@azure/msal-browser");

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
  const policyList = document.getElementById("policy-list")
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
          const wordFiles = await fetchWordFiles(accessToken);
          console.log("Word files:", wordFiles);
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

async function fetchWordFiles(accessToken) {

  // const documentUrl = Office.context.document.url;
  // const documentIdMatch = documentUrl.match(/sourcedoc=\{([^}]+)\}/);
  // const documentId = documentIdMatch ? documentIdMatch[1] : null;
  // console.log(documentId);

  

  // const endpoint = "https://graph.microsoft.com/v1.0/me/drive/root/search(q='.docx')";
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
  }
}

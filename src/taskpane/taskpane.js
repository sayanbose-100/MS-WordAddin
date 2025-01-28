const { PublicClientApplication } = require("@azure/msal-browser");

let policies = [];
let selectedPolicyId = "";
export let userEmail = "";

// const fetchPolicies = async (userEmail) => {
//   const policyContainer = document.getElementById("policies_container");

//   const response = await fetch(`http://localhost:3001/policies/${userEmail}`,{
//     method: "GET",
//     headers: {
//       "Content-Type": "application/json",
//     },
//   });

//   if (!response.ok) {
//     throw new Error(`Error fetching policies: ${response.statusText}`);
//   }
//   const data = await response.json();
//   if(data){
//     data.map((ele) => policies.push(ele));
//   }
//   console.log("Policies: ",policies);
//   // if(policies.length == 1) {
//   //   selectedPolicyId = policies[0].policyId;
//   //   console.log("Selected Policy: ", selectedPolicyId);
//   // }
//   if(policies.length > 0) {
//     const select = document.createElement("select");
//     select.id = "policies";
//     select.title = "policies";

//     policies.forEach((policy) => {
//     const option = document.createElement("option");
//     option.value = `${policy.policyId}`;
//     option.innerText = `${policy.policyName}`;
//     select.appendChild(option);
//     // console.log(policy);
//     })
//     policyContainer.appendChild(select);
//   }
// }

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
    document.getElementById("linkAcc").onclick = LinkAccount;
    document.getElementById("syncNow").onclick = saveDocumentCredentials;
  }
});

// export async function create() {
//   return Word.run(async (context) => {
//     const policyOwner = document.getElementById("policyOwner");
//     const policyName = document.getElementById("policyName");

//     policyOwner.style.fontSize = "20px";
//     policyName.style.color = "blue";

//     // policy entry
//     const policyEntry = `Name of the Owner: ${policyOwner.value} \\n Name of the Policy: ${policyName.value}`;

//     // Insert a paragraph at the end of the document.
//     const paragraph = context.document.body.insertParagraph(`${policyEntry}`, Word.InsertLocation.end);

//     await context.sync();
//   });
// }

// export async function syncF() {
//   const syncMessage = document.getElementById("sync_message");
//   const policyList = document.getElementById("policy-list");
//   initializeMsal()
//     .then(() => {
//       return Word.run(async (context) => {
//         try {
//           const loginResponse = await msalInstance.loginPopup({
//             scopes: ["Files.ReadWrite.All"],
//           });

//           const account = msalInstance.getAccount(loginResponse.account.username);
//           msalInstance.setActiveAccount(account);

//           const tokenResponse = await msalInstance.acquireTokenSilent({
//             scopes: ["Files.ReadWrite.All"],
//             account: account,
//           });

//           const accessToken = tokenResponse.accessToken;

//           // Use the access token to access OneDrive
//           console.log("Access token acquired:", accessToken);
//           syncMessage.innerText = "Token acquired";
//           syncMessage.style.color = "green";

//           // Fetch all MS Word files from the user's OneDrive
//           const wordFiles = await fetchCurrentWordFiles(accessToken);
//           console.log("Word files:", wordFiles);
//           saveDocumentCredentials(wordFiles);
//           console.log("Sent the files to the Server");
//         } catch (error) {
//           console.error("Authentication error:", error);
//         }
//       });
//     })
//     .catch(() => {
//       console.error("MSAL initialization failed");
//     });
// }

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
  return {
    id: data.value[0].id,
    webUrl: data.value[0].webUrl,
  };
}

async function saveDocumentCredentials() {
  const links = await createShareableLinkForCurrentFile();
  console.log(links);
  try {
    const payLoad = {
      editLink: links.editLink,
      readLink: links.readLink,
    }
    const response = await fetch(`http://localhost:4001/api/policies/policy/${selectedPolicyId}`, {
      method: "PATCH",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payLoad),
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

async function LinkAccount(accessToken) {
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
            const mail = data.mail || data.userPrincipalName;
            userEmail = mail;
            console.log("Mail ID: ",userEmail);
            // createShareableLinkForCurrentFile();
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

// graph API usage to generate a shareable link
async function generateShareableLink(itemId,type) {
  try {
    const tokenResponse = await msalInstance.acquireTokenSilent({
      scopes: ["Files.ReadWrite.All"],
      account: msalInstance.getActiveAccount(),
    });

    console.log(itemId);

    const accessToken = tokenResponse.accessToken;
    // console.log(accessToken);

    const endpoint = `https://graph.microsoft.com/v1.0/me/drive/items/${itemId}/createLink`;

    const response = await fetch(endpoint, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        type: `${type}`, // or "edit" depending on the type of link you want
        scope: "organization", // or "organization" depending on the scope of the link
      }),
    });

    if (!response.ok) {
      throw new Error(`Error creating shareable link: ${response.statusText}`);
    }

    const data = await response.json();
    const shareableLink = data.link.webUrl;
    console.log("Shareable Link: ", shareableLink);
    return `${shareableLink}?`;
  } catch (error) {
    console.error("Error generating shareable link:", error);
  }
}

function getCurrentFileName() {
  return new Promise((resolve, reject) => {
    Office.context.document.getFilePropertiesAsync((asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        const itemId = asyncResult.value.url.split('/').pop(); // Extract item_id from URL 
        resolve(itemId);
      } else {
        reject(new Error("Failed to get file properties"));
      }
    });
  });
}

async function createShareableLinkForCurrentFile() {
  try {
    const itemName = await getCurrentFileName();
    const itemId = await getItemId(itemName);
    const editLink = await generateShareableLink(itemId,"edit");
    const readLink = await generateShareableLink(itemId,"view");

    const payload = {
      editLink: editLink,
      readLink: readLink,
    }
    return payload;
  } catch (error) {
    console.error("Error creating shareable link for current file:", error);
  }
}

async function getItemId(itemName) {
  const endpoint = `https://graph.microsoft.com/v1.0/me/drive/root/search(q='${itemName}')?select=id,name,webUrl`;

  const tokenResponse = await msalInstance.acquireTokenSilent({
    scopes: ["Files.ReadWrite.All"],
    account: msalInstance.getActiveAccount(),
  });

  const accessToken = tokenResponse.accessToken;

  const response = await fetch(endpoint, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
  });

  if (!response.ok) {
    throw new Error(`Error fetching item ID: ${response.statusText}`);
  }

  const data = await response.json();
  return data.value[0].id;
}

const fetchPolicies = async () => {
  const policyContainer = document.getElementById("policyContainer");

  const response = await fetch(`http://localhost:4001/api/users/user/email/${userEmail}`,{
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
    const option = document.createElement("option");
    option.value = "All";
    option.innerText = "All";
    select.appendChild(option);

    policies.forEach((policy) => {
    const option = document.createElement("option");
    option.value = `${policy.policyId}`;
    option.innerText = `${policy.policyName}`;
    select.appendChild(option);
    })

    select.addEventListener("change", (event) => {
      selectedPolicyId = event.target.value;
      console.log("Selected Policy: ", selectedPolicyId);
    });

    policyContainer.appendChild(select);
  }
}
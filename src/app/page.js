"use client";

import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";
import {
  PublicClientApplication,
  InteractionRequiredAuthError,
} from "@azure/msal-browser";
import { useEffect, useState } from "react";
import { useRouter, useSearchParams } from "next/navigation";
import Image from "next/image";

const msalConfig = {
  auth: {
    clientId: process.env.AZURE_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}`,
    redirectUri: process.env.AZURE_REDIRECT_URI || "http://localhost:3001",
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false,
  },
};

let pca = new PublicClientApplication(msalConfig);

async function getAccessToken() {
  try {
    const accounts = pca.getAllAccounts();
    if (accounts.length > 0) {
      const silentRequest = {
        account: accounts[0],
        scopes: ["User.Read", "Files.Read.All"],
      };

      try {
        const response = await pca.acquireTokenSilent(silentRequest);
        console.log("アクセストークン (サイレント):", response.accessToken);
        return response.accessToken;
      } catch (error) {
        console.error("Silent token acquisition failed", error);
        if (error instanceof InteractionRequiredAuthError) {
          try {
            const response = await pca.acquireTokenPopup({
              scopes: ["User.Read", "Files.Read.All"],
            });
            console.log(
              "アクセストークン (ポップアップ):",
              response.accessToken
            );
            return response.accessToken;
          } catch (popupError) {
            console.error("Popup token acquisition failed", popupError);
          }
        }
      }
    }
    return null;
  } catch (error) {
    console.error("Something went wrong", error);
    return null;
  }
}

async function getOneDriveFiles(accessToken) {
  if (!accessToken) {
    return [];
  }

  const client = Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    },
  });

  const files = await client.api("/me/drive/root/children").get();
  console.log("APIレスポンス:", fil);
  console.log("valueの中身:", fil.value);
  return files.value;
}

export default function Home() {
  const [files, setFiles] = useState([]);
  const [isLoggedIn, setIsLoggedIn] = useState(false);
  const router = useRouter();
  const searchParams = useSearchParams();
  const code = searchParams.get("code");

  useEffect(() => {
    if (typeof window !== "undefined") {
      if (code) {
        async function loadFiles() {
          try {
            const accessToken = await getAccessToken();
            if (accessToken) {
              setIsLoggedIn(true);
              const oneDriveFiles = await getOneDriveFiles(accessToken);
              console.log("OneDrive ファイル:", oneDriveFiles);
              setFiles(oneDriveFiles);
            } else {
              setIsLoggedIn(false);
            }
          } catch (error) {
            console.error("loadFiles error", error);
            setIsLoggedIn(false);
          }
        }
        loadFiles();
      }
    }
  }, [code]);

  const handleSignIn = async () => {
    const clientId = process.env.AZURE_CLIENT_ID;
    const tenantId = process.env.AZURE_TENANT_ID;
    const redirectUri = "http://localhost:3001";
    const scope = "User.Read Files.Read.All";
    const state = "12345";

    const authUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize?client_id=${clientId}&response_type=code&redirect_uri=${redirectUri}&scope=${scope}&state=${state}`;
    window.location.href = authUrl;
  };

  return (
    <main className="flex min-h-screen flex-col items-center justify-between p-24">
      <h1>OneDrive Files</h1>
      {isLoggedIn ? (
        <>
          <p>OneDrive にログインしています</p>
          <button
            onClick={() => {
              pca.logoutRedirect({
                postLogoutRedirectUri: "http://localhost:3001",
              });
            }}
          >
            ログアウト
          </button>
          {files.length > 0 ? (
            <table>
              <thead>
                <tr>
                  <th>名前</th>
                  <th>種類</th>
                  <th>更新日時</th>
                </tr>
              </thead>
              <tbody>
                {files.map((file) => (
                  <tr key={file.id}>
                    <td>{file.name}</td>
                    <td>{file.fileSystemInfo?.fileType || "フォルダ"}</td>
                    <td>
                      {new Date(file.lastModifiedDateTime).toLocaleString()}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          ) : (
            <p>ファイルが見つかりませんでした。</p>
          )}
        </>
      ) : (
        <button onClick={handleSignIn}>OneDrive にログイン</button>
      )}
    </main>
  );
}

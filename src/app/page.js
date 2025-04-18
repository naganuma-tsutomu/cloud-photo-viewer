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
    clientId: process.env.NEXT_PUBLIC_AZURE_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.NEXT_PUBLIC_AZURE_TENANT_ID}`,
    redirectUri: process.env.NEXT_PUBLIC_AZURE_REDIRECT_URI || "http://localhost:3000",
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false,
  },
};

let pca = new PublicClientApplication(msalConfig);

async function initializePca() {
  await pca.initialize();
}

async function getAccessToken(code) {
  if (code) {
    try {
      const response = await fetch('/api/token', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ code }),
      });

      const data = await response.json();
      if (response.ok) {
        console.log("アクセストークン (コード):", data.access_token);
        return data.access_token;
      } else {
        console.error("Failed to retrieve access token", data);
        return null;
      }
    } catch (error) {
      console.error("Error exchanging code for token", error);
      return null;
    }
  } else {
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
                "アクセストークン (ポップアップ):", response.accessToken
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
  }

async function getOneDriveFiles(accessToken, path) {
  if (!accessToken) {
    return [];
  }

  const client = Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    },
  });

  const apiPath = path === "/" ? "/me/drive/root/children" : `/me/drive/root:${path}:/children`;
  const files = await client.api(apiPath).get();
  console.log("APIレスポンス:", files);
  console.log("valueの中身:", files.value);
  return files.value;
}

function getParentPath(path) {
  if (path === "/") {
    return null;
  }
  const parts = path.split("/");
  parts.pop();
  return parts.length > 1 ? parts.join("/") : "/";
}

export default function Home() {
  const [files, setFiles] = useState([]);
  const [isLoggedIn, setIsLoggedIn] = useState(false);
  const [accessToken, setAccessToken] = useState(null);
  const router = useRouter();
  const searchParams = useSearchParams();
  const code = searchParams.get("code");
  const dir = searchParams.get("dir") || "/";

  useEffect(() => {
    async function initializeAndLoadFiles() {
      await initializePca();
      if (typeof window !== "undefined") {
        // localStorageからログイン状態を読み込む
        const storedIsLoggedIn = localStorage.getItem('isLoggedIn') === 'true';
        setIsLoggedIn(storedIsLoggedIn);

        if (code) {
          try {
            const token = await getAccessToken(code);
            console.log(token);
            if (token) {
              localStorage.setItem('accessToken', token); // アクセストークンをlocalStorageに保存
              localStorage.setItem('isLoggedIn', 'true'); // ログイン状態をlocalStorageに保存
              setIsLoggedIn(true);
              setAccessToken(token);
              const oneDriveFiles = await getOneDriveFiles(token, dir);
              console.log("OneDrive ファイル:", oneDriveFiles);
              setFiles(oneDriveFiles);
              router.replace('/'); // codeパラメータを削除
            } else {
              setIsLoggedIn(false);
            }
          } catch (error) {
            console.error("loadFiles error", error);
            setIsLoggedIn(false);
          }
        } else {
          // localStorageからアクセストークンを読み込む
          const storedAccessToken = localStorage.getItem('accessToken');
          if (storedAccessToken) {
            try {
              localStorage.setItem('isLoggedIn', 'true'); // ログイン状態をlocalStorageに保存
              setIsLoggedIn(true);
              setAccessToken(storedAccessToken);
              const oneDriveFiles = await getOneDriveFiles(storedAccessToken, dir);
              console.log("OneDrive ファイル (localStorage):", oneDriveFiles);
              setFiles(oneDriveFiles);
            } catch (error) {
              console.error("loadFiles error", error);
              setIsLoggedIn(false);
              localStorage.removeItem('accessToken'); // エラーが発生した場合はlocalStorageから削除
              localStorage.removeItem('isLoggedIn'); // ログイン状態をlocalStorageから削除
              // ログイン画面にリダイレクト
              router.push('/');
            }
          }
        }
      }
    }
    initializeAndLoadFiles();
  }, [code, router, setIsLoggedIn, dir]);

  const handleSignIn = async () => {
    const clientId = process.env.NEXT_PUBLIC_AZURE_CLIENT_ID;
    const tenantId = process.env.NEXT_PUBLIC_AZURE_TENANT_ID;
    const redirectUri = "http://localhost:3000";
    const scope = "User.Read Files.Read.All";
    const state = "12345";

    const authUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize?client_id=${clientId}&response_type=code&redirect_uri=${redirectUri}&scope=${scope}&state=${state}`;
    window.location.href = authUrl;
  };

  const parentPath = getParentPath(dir);

  const imageFiles = files.filter(file => file.file && file.file.mimeType.startsWith('image/'));

  return (
    <main className="flex min-h-screen flex-col items-center justify-between p-24">
      <h1 className="text-2xl font-bold mb-4">OneDrive Files</h1>
      {isLoggedIn ? (
        <>
          <p className="mb-2">OneDrive にログインしています</p>
          <button
            className="bg-blue-500 hover:bg-blue-700 text-white font-bold py-2 px-4 rounded mb-4"
            onClick={() => {
              localStorage.removeItem('accessToken'); // アクセストークンをlocalStorageから削除
              localStorage.removeItem('isLoggedIn'); // ログイン状態をlocalStorageから削除
              if (pca) {
                pca.logoutRedirect({
                  postLogoutRedirectUri: "http://localhost:3000",
                });
              }
            }}
          >
            ログアウト
          </button>
          {parentPath && (
            <a
              className="text-blue-500 hover:text-blue-700 mb-4"
              href={`/?dir=${parentPath}`}
            >
              上の階層へ
            </a>
          )}
          {imageFiles.length > 0 ? (
            <div className="flex flex-wrap justify-center">
              {imageFiles.map(file => (
                <div key={file.id} className="m-2">
                  <Image
                    src={`https://graph.microsoft.com/v1.0/me/drive/items/${file.id}/content?access_token=${accessToken}`}
                    alt={file.name}
                    width={200}
                    height={200}
                    onError={(e) => {
                      e.target.onerror = null;
                      e.target.src = "/file.svg";
                    }}
                  />
                </div>
              ))}
            </div>
          ) : (
            <>
              <p>ファイルが見つかりませんでした。</p>
              <table className="table-auto">
                <thead>
                  <tr className="bg-gray-200">
                    <th className="px-4 py-2">名前</th>
                    <th className="px-4 py-2">種類</th>
                    <th className="px-4 py-2">更新日時</th>
                  </tr>
                </thead>
                <tbody>
                  {files.map((file) => (
                    <tr key={file.id}>
                      <td className="border px-4 py-2">
                        {file.folder ? (
                          <a
                            className="text-blue-500 hover:text-blue-700"
                            href={`/?dir=${dir === "/" ? "/" + file.name : dir + "/" + file.name}`}
                          >
                            {file.name}
                          </a>
                        ) : (
                          file.name
                        )}
                      </td>
                      <td className="border px-4 py-2">
                        {file.fileSystemInfo?.fileType || (file.folder ? "フォルダ" : "ファイル")}
                      </td>
                      <td className="border px-4 py-2">
                        {new Date(file.lastModifiedDateTime).toLocaleString()}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </>
          )}
        </>
      ) : (
        <button className="bg-blue-500 hover:bg-blue-700 text-white font-bold py-2 px-4 rounded" onClick={handleSignIn}>OneDrive にログイン</button>
      )}
    </main>
  );
}

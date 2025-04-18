import { NextResponse } from 'next/server';

export async function POST(request) {
  const { code } = await request.json();

  const tokenEndpoint = `https://login.microsoftonline.com/${process.env.NEXT_PUBLIC_AZURE_TENANT_ID}/oauth2/v2.0/token`;
  const params = new URLSearchParams();
  params.append("client_id", process.env.NEXT_PUBLIC_AZURE_CLIENT_ID);
  params.append("scope", "User.Read Files.Read.All");
  params.append("code", code);
  params.append("redirect_uri", process.env.NEXT_PUBLIC_AZURE_REDIRECT_URI || "http://localhost:3000");
  params.append("grant_type", "authorization_code");
  params.append("client_secret", process.env.NEXT_PUBLIC_AZURE_CLIENT_SECRET);

  try {
    const response = await fetch(tokenEndpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
      body: params.toString(),
    });

    console.log("APIリクエスト:", {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
      body: params.toString(),
    }); // 追加

    const data = await response.json();
    console.log("APIレスポンス:", data); // 追加
    if (response.ok) {
      console.log("アクセストークン (コード):", data.access_token);
      return NextResponse.json({ access_token: data.access_token });
    } else {
      console.error("Failed to retrieve access token", data);
      return NextResponse.json({ error: "Failed to retrieve access token" }, { status: 500 });
    }
  } catch (error) {
    console.error("Error exchanging code for token", error);
    return NextResponse.json({ error: "Error exchanging code for token" }, { status: 500 });
  }
}

import axios from "axios";

export default async function accessToken() {
  try {
    var data = new FormData();
    data.append("grant_type", "client_credentials");
    data.append("client_id", "faaa875a-b636-475c-b0f5-85f76dcd5208@bea1b417-4237-40b8-b020-57fce9abdb43");
    data.append("client_secret", "CxYMGEA1sNjj5G3SMzWAx6cA0dGXjcXWr+qf03u+kFU=");
    data.append(
      "resource",
      "00000003-0000-0ff1-ce00-000000000000/salic.sharepoint.com@bea1b417-4237-40b8-b020-57fce9abdb43"
    );

    let response = await axios({
      method: "POST",
      url: `https://accounts.accesscontrol.windows.net/bea1b417-4237-40b8-b020-57fce9abdb43/tokens/oAuth/2`,
      data: data,
    });
    return response;
  } catch (err) {
    console.log(err.response);
    return err.response;
  }
}

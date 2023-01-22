import axios from "axios";

export default async function getDraftReplys(email, token) {
  try {
    let response = await axios({
      method: "GET",
      url: `https://salic.sharepoint.com/sites/portal/_api/web/lists/getbytitle('Outlook Draft Replys')/items?$filter=To eq '${email}'&$orderby=Created desc`,
      headers: {
        Accept: "application/json",
        Authorization: `Bearer ${token}`,
      },
    });
    return response;
  } catch (err) {
    console.log(err.response);
    return err.response;
  }
}

import axios from "axios";

export default async function addDraftReply(token, data) {
  try {
    let response = await axios({
      method: "POST",
      url: "https://salic.sharepoint.com/sites/portal/_api/web/lists/getbytitle('Outlook Draft Replys')/items",
      headers: {
        Accept: "application/json",
        Authorization: `Bearer ${token}`,
      },
      data: data,
    });
    return response;
  } catch (err) {
    console.log(err.response);
    return err.response;
  }
}

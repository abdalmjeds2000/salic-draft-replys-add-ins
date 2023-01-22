import React, { useEffect } from "react";
import { Label, MessageBar, MessageBarType, PrimaryButton, Spinner, TextField, TooltipHost } from "@fluentui/react";
import addDraftReply from "../api/addDraftReply";
import axios from "axios";
import AutoCompleteUsers from "./AutoCompleteUsers";

const DraftReply = (props) => {
  const [toField, setToField] = React.useState("");
  const [toFieldError, setToFieldError] = React.useState(false);
  const [notesField, setNotesField] = React.useState("");
  const [notesFieldError, setNotesFieldError] = React.useState(false);
  const [loading, setLoading] = React.useState(true);
  const [isShow, setIsShow] = React.useState(false);
  const [count, setCount] = React.useState(0);
  const [showSuccessMessage, setShowSuccessMessage] = React.useState(false);

  const token =
    "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6Ii1LSTNROW5OUjdiUm9meG1lWm9YcWJIWkdldyIsImtpZCI6Ii1LSTNROW5OUjdiUm9meG1lWm9YcWJIWkdldyJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvc2FsaWMuc2hhcmVwb2ludC5jb21AYmVhMWI0MTctNDIzNy00MGI4LWIwMjAtNTdmY2U5YWJkYjQzIiwiaXNzIjoiMDAwMDAwMDEtMDAwMC0wMDAwLWMwMDAtMDAwMDAwMDAwMDAwQGJlYTFiNDE3LTQyMzctNDBiOC1iMDIwLTU3ZmNlOWFiZGI0MyIsImlhdCI6MTY3NDM3MTMwNiwibmJmIjoxNjc0MzcxMzA2LCJleHAiOjE2NzQ0NTgwMDYsImlkZW50aXR5cHJvdmlkZXIiOiIwMDAwMDAwMS0wMDAwLTAwMDAtYzAwMC0wMDAwMDAwMDAwMDBAYmVhMWI0MTctNDIzNy00MGI4LWIwMjAtNTdmY2U5YWJkYjQzIiwibmFtZWlkIjoiZmFhYTg3NWEtYjYzNi00NzVjLWIwZjUtODVmNzZkY2Q1MjA4QGJlYTFiNDE3LTQyMzctNDBiOC1iMDIwLTU3ZmNlOWFiZGI0MyIsIm9pZCI6IjdhNmEyNmYxLTAzZTgtNDBiZC1iMDI2LTk5ODdkNjkxMWIyMiIsInN1YiI6IjdhNmEyNmYxLTAzZTgtNDBiZC1iMDI2LTk5ODdkNjkxMWIyMiIsInRydXN0ZWRmb3JkZWxlZ2F0aW9uIjoiZmFsc2UifQ.FmnpGaR-XmH_OxFDBpTRnXApKEwENt2nrtBzXhYmZz_hqhp-7KR7GMPE5TUMHc9_s7yeTYK3D7g_Fhkt1ZcWzBkGPjBxPM60g_2OQ_90bbv_ubKteM6bA25UDL7JjMnJLJhGvezI0WtjEGdJZpVCAusxvdinPDtwqLsyNGKYsiOPAw-n0_tZ8FjjDkDMmr6KKg3B2N3lmMvcZ0N5VJvxp7oMdS62qqj7Au_CQB9RmdlX8xCb5cUjS_LE0O51-56Qprh8y2dfGVeGwxc_0Q5wNhRRKnv9RXWGQN_Q12pkP6nTveg7Zoy5MeqefhVCfcW8lA_yUegZq1GrWP3w1NP_Ig";

  const handleSubmit = async (body) => {
    if (toField.trim() === "") {
      setToFieldError(true);
    } else {
      setToFieldError(false);
    }
    if (notesField.trim() === "") {
      setNotesFieldError(true);
    } else {
      setNotesFieldError(false);
    }
    if (toField.trim() !== "" && notesField.trim() !== "") {
      const payload = {
        Title: Office.context.mailbox.item.subject,
        MailId: Office.context.mailbox.item.itemId,
        From: Office.context.mailbox.item.from.emailAddress,
        MailDate: new Date(Office.context.mailbox.item.dateTimeCreated).toISOString(),
        Body: body.html,
        BodyText: body.text,
        To: toField.trim(),
        Notes: notesField.trim(),
      };

      setLoading(true);
      const item = await addDraftReply(token, payload);
      if (item.status == 201) {
        setShowSuccessMessage(true);
        setToField("");
        setNotesField("");
        setCount((prev) => prev + 1);
      }
      setLoading(false);
    }
  };
  const prepareMailBody = (text) => {
    Office.context.mailbox.item.body.getAsync(
      "html",
      { asyncContext: "This is passed to the callback" },
      function callback(result) {
        handleSubmit({ html: result.value, text: text });
      }
    );
  };
  const prepareMailText = () => {
    Office.context.mailbox.item.body.getAsync(
      "text",
      { asyncContext: "This is passed to the callback" },
      function callback(result) {
        prepareMailBody(result.value);
      }
    );
  };

  const checkMail = async (mailId) => {
    try {
      let response = await axios({
        method: "GET",
        url: `https://salic.sharepoint.com/sites/portal/_api/web/lists/getbytitle('Outlook Draft Replys')/items?$filter=MailId eq '${mailId}'`,
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
  };
  useEffect(async () => {
    setLoading(true);
    const mailId = Office.context.mailbox.item.itemId;
    const response = await checkMail(mailId);
    if (response.status == 200) {
      if (response.data.value.length == 0) {
        setIsShow(true);
      } else {
        setIsShow(false);
      }
    }
    setLoading(false);
  }, [count]);

  if (loading) {
    return (
      <div style={{ margin: "50px 0" }}>
        <Spinner label="Wait..." ariaLive="assertive" labelPosition="right" />
      </div>
    );
  }
  if (!isShow) {
    return (
      <div style={{ padding: 25 }}>
        {showSuccessMessage ? <SuccessMessage /> : <MessageBar>This Email is Drafted.</MessageBar>}
      </div>
    );
  }

  const checkIcon = { iconName: "CheckMark" };

  return (
    <main className="ms-welcome__main">
      <h2 className="ms-font-xl ms-fontWeight-semilight ms-fontColor-neutralPrimary ms-u-slideUpIn20">
        {props.message}
      </h2>

      <div style={{ marginBottom: 15 }}>
        {/* <Label htmlFor="toUserField" required>
          Email Address
        </Label>
        <TextField
          id="toUserField"
          type="email"
          placeholder="e.g. ahmed@outlook.com"
          value={toField}
          onChange={(e) => setToField(e.target.value)}
          errorMessage={toFieldError}
        /> */}
        <AutoCompleteUsers setSelectedUser={setToField} selectedUser={toField} fieldError={toFieldError} />

        <Label htmlFor="notes" required>
          Notes
        </Label>
        <TextField
          id="notes"
          multiline
          placeholder="Write some notes"
          rows={6}
          value={notesField}
          onChange={(e) => setNotesField(e.target.value)}
          errorMessage={notesFieldError}
        />
      </div>

      <TooltipHost content="Click Here to Ask for Reply">
        <PrimaryButton
          iconProps={checkIcon}
          text="Ask for Reply"
          onClick={prepareMailText}
          allowDisabledFocus
          disabled={loading}
        />
      </TooltipHost>
    </main>
  );
};

export default DraftReply;

const SuccessMessage = () => (
  <MessageBar messageBarType={MessageBarType.success} isMultiline={false} dismissButtonAriaLabel="Close">
    Your Action has been done Successfully.
  </MessageBar>
);

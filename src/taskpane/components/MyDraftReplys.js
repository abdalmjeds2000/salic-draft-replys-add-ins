import { MessageBar, Spinner, Text, TooltipHost } from "@fluentui/react";
import React, { useState, useEffect } from "react";
import getDraftReplys from "../api/getDraftReplys";
import PreviewItem from "./PreviewItem";

const MyDraftReplys = () => {
  const [mailsList, setMailsList] = useState([]);
  const [mailContent, setMailContent] = useState({});
  const [loading, setLoading] = useState(false);
  const [showItem, setShowItem] = useState(false);

  const fetchItems = async () => {
    setLoading(true);
    const email = Office.context.mailbox.userProfile.emailAddress;
    const token =
      "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6Ii1LSTNROW5OUjdiUm9meG1lWm9YcWJIWkdldyIsImtpZCI6Ii1LSTNROW5OUjdiUm9meG1lWm9YcWJIWkdldyJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvc2FsaWMuc2hhcmVwb2ludC5jb21AYmVhMWI0MTctNDIzNy00MGI4LWIwMjAtNTdmY2U5YWJkYjQzIiwiaXNzIjoiMDAwMDAwMDEtMDAwMC0wMDAwLWMwMDAtMDAwMDAwMDAwMDAwQGJlYTFiNDE3LTQyMzctNDBiOC1iMDIwLTU3ZmNlOWFiZGI0MyIsImlhdCI6MTY3NDM3MTMwNiwibmJmIjoxNjc0MzcxMzA2LCJleHAiOjE2NzQ0NTgwMDYsImlkZW50aXR5cHJvdmlkZXIiOiIwMDAwMDAwMS0wMDAwLTAwMDAtYzAwMC0wMDAwMDAwMDAwMDBAYmVhMWI0MTctNDIzNy00MGI4LWIwMjAtNTdmY2U5YWJkYjQzIiwibmFtZWlkIjoiZmFhYTg3NWEtYjYzNi00NzVjLWIwZjUtODVmNzZkY2Q1MjA4QGJlYTFiNDE3LTQyMzctNDBiOC1iMDIwLTU3ZmNlOWFiZGI0MyIsIm9pZCI6IjdhNmEyNmYxLTAzZTgtNDBiZC1iMDI2LTk5ODdkNjkxMWIyMiIsInN1YiI6IjdhNmEyNmYxLTAzZTgtNDBiZC1iMDI2LTk5ODdkNjkxMWIyMiIsInRydXN0ZWRmb3JkZWxlZ2F0aW9uIjoiZmFsc2UifQ.FmnpGaR-XmH_OxFDBpTRnXApKEwENt2nrtBzXhYmZz_hqhp-7KR7GMPE5TUMHc9_s7yeTYK3D7g_Fhkt1ZcWzBkGPjBxPM60g_2OQ_90bbv_ubKteM6bA25UDL7JjMnJLJhGvezI0WtjEGdJZpVCAusxvdinPDtwqLsyNGKYsiOPAw-n0_tZ8FjjDkDMmr6KKg3B2N3lmMvcZ0N5VJvxp7oMdS62qqj7Au_CQB9RmdlX8xCb5cUjS_LE0O51-56Qprh8y2dfGVeGwxc_0Q5wNhRRKnv9RXWGQN_Q12pkP6nTveg7Zoy5MeqefhVCfcW8lA_yUegZq1GrWP3w1NP_Ig";
    const items = await getDraftReplys(email, token);
    if (items.status == 200) {
      console.log(items);
      setMailsList(items.data.value);
    }
    setLoading(false);
  };

  useEffect(() => {
    fetchItems();
  }, []);

  const handleOpenItem = (item) => {
    setMailContent(item);
    setShowItem(true);
  };

  const handleCloseItem = () => {
    setMailContent({});
    setShowItem(false);
  };

  if (loading) {
    return (
      <div style={{ margin: "50px 0" }}>
        <Spinner label="Loading..." ariaLive="assertive" labelPosition="right" />
      </div>
    );
  }
  if (showItem) {
    return <PreviewItem item={mailContent} handleClose={handleCloseItem} />;
  }
  return (
    <div className="my-draft-replys">
      <div className="header"></div>
      <div className="items">
        {mailsList.length > 0 ? (
          mailsList.map((item, i) => (
            <div key={i} className="mail-box">
              <Text variant="xLarge" block className="title">
                <TooltipHost content={`Mail Date: ${new Date(item.MailDate).toLocaleString("en-US")}`}>
                  <a onClick={() => handleOpenItem(item)}>{item.Title}</a>
                </TooltipHost>
              </Text>
              <Text block style={{ color: "#777", marginBottom: 10 }}>
                Drafted at {new Date(item.Created).toLocaleString("en-US")}
              </Text>

              <Text variant="small" block style={{ lineHeight: "15px", height: 45, overflow: "hidden" }}>
                <div dangerouslySetInnerHTML={{ __html: item.BodyText }}></div>
              </Text>
            </div>
          ))
        ) : (
          <MessageBar>There is no emails drafted for you.</MessageBar>
        )}
      </div>
      <div className="footer"></div>
    </div>
  );
};

export default MyDraftReplys;

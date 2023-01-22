import React from "react";
import { DefaultButton, Popup, PrimaryButton, Text, TooltipHost } from "@fluentui/react";
import { useBoolean } from "@fluentui/react-hooks";

const styles = {
  mailContainer: {
    padding: 25,
  },
  header: {
    marginBottom: 15,
  },
  subject: {
    marginBottom: 5,
  },
  date: {
    marginBottom: 25,
    color: "#777",
  },
  notesTitle: {
    marginBottom: 10,
  },
  notes: {
    borderRadius: "0 15px 15px 15px",
    padding: "25px 22px",
    backgroundColor: "#eee",
    marginBottom: 25,
    fontSize: "1.2rem",
  },
};
const expandIcon = { iconName: "ChevronRightSmall" };
const collapseIcon = { iconName: "ChevronDownSmall" };
const backIcon = { iconName: "ChromeBack" };

const PreviewItem = ({ handleClose, item }) => {
  const [isPopupVisible, { toggle: toggleIsPopupVisible }] = useBoolean(false);

  return (
    <div style={styles.mailContainer}>
      <div style={styles.header}>
        <DefaultButton text="Back To My Draft Replys" iconProps={backIcon} onClick={handleClose} allowDisabledFocus />
      </div>

      <div>
        <div>
          <Text variant={"xxLarge"} block style={styles.subject}>
            <TooltipHost content={`Mail Date: ${new Date(item.MailDate).toLocaleString("en-US")}`}>
              {item.Title}
            </TooltipHost>
          </Text>
          <span></span>
          <Text variant={"medium"} nowrap block style={styles.date}>
            Drafted at {new Date(item.Created).toLocaleString("en-US")}
          </Text>
        </div>
        <div style={{ marginBottom: 25 }}>
          <PrimaryButton
            iconProps={isPopupVisible ? collapseIcon : expandIcon}
            onClick={toggleIsPopupVisible}
            text="Notes"
          />
          {isPopupVisible && (
            <Popup>
              <div dangerouslySetInnerHTML={{ __html: item.Notes }} style={styles.notes}></div>
            </Popup>
          )}
        </div>

        <div>
          <Text variant="xLarge" block style={{ backgroundColor: "#eee", padding: 20, borderRadius: "15px 15px 0 0" }}>
            Mail Preview
          </Text>
          <div style={{ backgroundColor: "#f7f7f7", overflow: "auto" }}>
            <div dangerouslySetInnerHTML={{ __html: item.Body }}></div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default PreviewItem;

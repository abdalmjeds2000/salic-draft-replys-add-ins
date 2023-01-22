import React, { useState } from "react";
import { ChoiceGroup, Label, Text, TextField } from "@fluentui/react";
import axios from "axios";

let CancelToken = axios.CancelToken;

const AutoCompleteUsers = ({ setSelectedUser, selectedUser, fieldError }) => {
  const [users, setUsers] = useState([]);
  const [loading, setLoading] = useState(false);

  let cancel;
  const handleChange = async (e) => {
    if (cancel) {
      cancel();
    }
    const q = e.target.value;
    if (q.length >= 3) {
      setLoading(true);
      const response = await axios.get(
        `https://salicapi.com/api/User/AutoComplete?term=${q}&_type=query&q=${q}&_=1667805757891`,
        { cancelToken: new CancelToken((c) => (cancel = c)) }
      );
      if (response?.status == 200 && response?.data?.Status == 200) {
        const result = response.data.Data.value || [];
        if (result.length > 0) {
          setUsers(result);
        } else {
          setUsers([]);
          setSelectedUser("");
        }
      }
      setLoading(false);
    } else if (q.length === 0) {
      setUsers([]);
      setSelectedUser("");
    }
  };

  const options = users.map((item) => ({
    key: item.mail,
    text: <UserItem {...item} />,
  }));

  return (
    <div>
      <Label htmlFor="toUserField" required>
        Email Address
      </Label>
      <TextField
        id="toUserField"
        placeholder="username@salic.com"
        onChange={handleChange}
        suffix={loading ? "loading..." : "Type to Search"}
        errorMessage={fieldError && selectedUser === "" ? "Please Pick User first" : ""}
      />

      {users.length > 0 && (
        <div style={{ marginBottom: 15 }}>
          <ChoiceGroup label="Pick one User" options={options} onChange={(_e, option) => setSelectedUser(option.key)} />
        </div>
      )}
    </div>
  );
};

export default AutoCompleteUsers;

const UserItem = (item) => (
  <div style={{ textAlign: "left" }}>
    <Text block variant="xLargePlus" style={{ fontSize: "1em" }}>
      {item.displayName}
    </Text>
    <Text block style={{ fontSize: ".8em" }}>
      {item.mail}
    </Text>
  </div>
);

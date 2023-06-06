import fetch from "node-fetch";
import {promiseHandler} from "./util.js";

const {
    FETCH_URL,
} = process.env;

export const request = async body => {
    const [requestError, response] = await promiseHandler(fetch(FETCH_URL, {
        method: "POST",
        body,
        headers: {
            "Content-Type": "application/x-www-form-urlencoded",
        },
    }));
    if (requestError) return console.error(requestError);

    const [jsonParseError, object] = await promiseHandler(response.json());
    if (jsonParseError) return console.error(jsonParseError);

    return object;
}
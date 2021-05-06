# Outlook email processor add-in
This add-in allows the user to extract specific fields from an email based on configurable JSON templates. Furthermore, URL's in the shape of an action specific to the configured templates can also be extracted, thus enabling the user to send a chain of HTTP requests related to the starting URL found in the email body.

# Prerequisites
```shell
> nodejs : https://nodejs.org/en/
> npm : https://www.npmjs.com/get-npm
> Outlook Web / Desktop 2013
```

# Setup
```shell
> Clone/Download the repository.
> Run `npm install` to install npm dependencies.
> Outlook Web:
>   Select a random message.
>   Press `More actions` at the top of the message.
>   Press `Get Add-Ins` at the bottom of the list.
>   Choose `My Add-Ins` on the left panel.
>   Press `Add a custom add-in` at the buttom of the dialog.
>   Choose `Add from file` and select the `manifest.xml` file from the add-in folder.
> Outlook Desktop:
>   Select the `Home` column.
>   Press the `Get Add-Ins` button.
>   Choose `My Add-Ins` on the left panel.
>   Press `Add a custom add-in` at the buttom of the dialog.
>   Choose `Add from file` and select the `manifest.xml` file from the add-in folder.
```

# Configuration
The main file used for configuration is `patterns.json`, located in `src/taskpane`. This JSON file is structured as follows:
```json
{
    "patterns": [
        {
            "Test pattern {email}": {
                "description": "Test description.",
                "actions": [
                    {
                        "Test action - {url}": {
                            "requests": [
                                {
                                    "url": "https://localhost:3000/firstRequest.html",
                                    "type": "GET"
                                },
                                {
                                    "url": "https://localhost:3000/secondRequest.html",
                                    "type": "PUT",
                                    "params": [
                                        "name=test"
                                    ]
                                },
                                {
                                    "url": "https://localhost:3000/thirdRequest.html",
                                    "type": "POST",
                                    "params": [
                                        "request2param=name"
                                    ]
                                }
                            ]
                        }
                    }
                ]
            },
            "Second pattern {user}": {
                "description": "Pattern without any actions defined."
            }
        }
    ]
}
```
The params array from each request also allows for 2 special string literals: `requestREQ_NOparam=PARAM_NAME` and `requestREQ_NObody=PARAM_NAME`, where REQ_NO is the index of the request we want to take information from, "param" and "body" are the locations where we want to look for data (request's used parameters & HTTP body response).

Based on the JSON structure example given previously, the workflow is as follows:
```shell
> We loop through each pattern defined in the JSON file, match the subject/body with the pattern\'s key: "Test pattern {email}" and extract special fields defined between curly brackets.
> We match the action key, "Test action - {url}" and extract its starting URL. After that we send the initial HTTP request, process the next request (special params that are based on the previously executed requests) defined in this action and send it. Requests are being sent sequentially until there\'s none left in the request path.
```

Depending on the user's needs, special fields can also be added/modified:
```shell
> Declare regex literal for that specific field inside `src/taskpane/taskpane.js`.
> Add it to the `regexMap` dictionary, inside `src/taskpane/taskpane.js` line 50.
> Add the field to the replace function, inside `src/taskpane/taskpane.js` line 76.
```
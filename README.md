# ms-registeration
Static website to for registration form that post the input to store in google sheet

## Editable config

`index.html` loads config from the Google Apps Script web app using JSONP, then falls back to `config.json` if the remote config cannot be loaded.

The Apps Script in `apps-script/Code.gs` writes to this spreadsheet:

`1AFff64g1PR6qtcy9YzNR6xl1GNRbXFg33t8XPm3YyJw`

To enable editing from `admin.html`:

1. Open the Google Apps Script project for the registration web app.
2. Add or merge the code from `apps-script/Code.gs`.
3. In Apps Script, set a script property named `ADMIN_TOKEN` to a private value.
4. Deploy the Apps Script as a web app.
5. Open `admin.html`, enter the same admin token, edit the config, and save.

Registrations are stored in the existing `formData` sheet.

Editable admin data is stored in these tabs:

- `Config`: event settings and form titles
- `MainQuestions`: main registration questions
- `QuestionnaireQuestions`: additional questionnaire questions

The admin page can load the default examples from `config.json`, `main.json`, and `questionnaire.json`, then save them into the spreadsheet tabs.

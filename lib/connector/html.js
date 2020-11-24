const providerName = process.env.vendor_name || 'OAuth';

exports.getOAuthCallbackPageHtml = (verificationCode) => `<html>
<head>
    <title>${providerName}</title>
    <meta name="viewport" content="width=device-width, initial-scale=1.0" >
</head>
<body>
    <div id="instructionText" style="display: none">
        <div class="instruction-text">
            <p>You're almost there!</p>
            <p>To finish linking your account with ${providerName}, type</p>
            <span class="verification-code">${verificationCode}</span>
            <p>in the Microsoft Teams chat window.</p>
        </div>
    </div>

    <script src="https://statics.teams.cdn.office.net/sdk/v1.6.0/js/MicrosoftTeams.min.js" integrity="sha384-mhp2E+BLMiZLe7rDIzj19WjgXJeI32NkPvrvvZBrMi5IvWup/1NUfS5xuYN5S3VT" crossorigin="anonymous"></script>
    <script type="text/javascript">
        // If the window is still visible after 5 seconds, then we are probably on a platform
        // that does not support automatically passing the verification code using notifySuccess().
        // So we ask the user to manually enter the verification code in the chat window.
        setTimeout(function () {
            document.getElementById("instructionText").style.display = "initial";
        }, 5000);
        microsoftTeams.initialize();
        microsoftTeams.authentication.notifySuccess("${verificationCode}");
    </script>
</body>
</html>`;

exports.getOAuthErrorPageHtml = (reason) => `<html>
<head>
    <title>${providerName}</title>
    <meta name="viewport" content="width=device-width, initial-scale=1.0" >
</head>
<body>
    <div>
        <div>
            <p>Authorization Error</p>
            <p>${reason}</p>
        </div>
    </div>
</body>
</html>`;

exports.getOAuthStartPageHtml = (authorizationUrlBase) => `<html>
<head>
    <title>Sign In</title>
</head>
<body>
    <script src="https://statics.teams.microsoft.com/sdk/v1.2/js/MicrosoftTeams.min.js" integrity="sha384-OncOcMprEYFkUvhe15zi8VdCZIOcuNGnf4ilq2yTfDEjTwPF19V7pZrzxOq/iqt0" crossorigin="anonymous"></script>
    <script type="text/javascript">
        microsoftTeams.initialize();

        // Parse query parameters
        let queryParams = {};
        location.search.substr(1).split("&").forEach(function(item) {
            let s = item.split("="),
            k = s[0],
            v = s[1] && decodeURIComponent(s[1]);
            queryParams[k] = v;
        });

        // Restrict to expected URLs only, so this page isn't used as a springboard to malicious sites
        function isValidAuthorizationUrl(url) {
            return url.indexOf("${authorizationUrlBase}") === 0;
        }

        let authorizationUrl = queryParams["authorizationUrl"];
        if (!authorizationUrl || !isValidAuthorizationUrl(authorizationUrl)) {
            microsoftTeams.authentication.notifyFailure("Invalid authorization url");
        } else {
            window.location.assign(authorizationUrl);
        }
    </script>
</body>
</html>`;

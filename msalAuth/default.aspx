<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="default.aspx.cs" Inherits="msalAuth._default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>

            <asp:Label ID="Label1" runat="server" Text="Loading..."></asp:Label>

            <script src="https://statics.teams.microsoft.com/sdk/v1.0/js/MicrosoftTeams.min.js"></script>
            <script src="https://ajax.googleapis.com/ajax/libs/jquery/2.1.3/jquery.min.js"></script>
            <script>

                microsoftTeams.initialize();

                var field = 'tk';

                var url = window.location.href;

                if (url.indexOf('?' + field + '=') == -1) {
                    var token = null;

                    if (token == null) {
                        microsoftTeams.authentication.authenticate({
                            url: "/auth.html",
                            width: 400,
                            height: 400,
                            successCallback: (data) => {
                                token = data;
                                window.location = 'default.aspx?tk=' + token;
                            },
                            failureCallback: function (err) {
                                document.getElementById("Label1").innerHTML = "Failed to authenticate and get token.<br/>" + err;
                            }
                        });
                    }
                    else {
                        window.location = 'default.aspx?tk=' + token;
                    }
                }
            </script>
        </div>
    </form>
</body>
</html>

$(document).ready(function () {
  $("#run").click(() => tryCatch(getKakuin));
});

function getKakuin() {
  var authenticator;
  var client_id = "2b588edd-4adc-4e4f-a3e6-a8d362246a29";
  var redirect_url = "https://mikiyks.github.io/inkan4/";
  var scope = "https://graph.microsoft.com/Files.Read.All";
  var access_token;

  authenticator = new OfficeHelpers.Authenticator();

  //access_token取得
  authenticator.endpoints.registerMicrosoftAuth(client_id, {
    redirectUrl: redirect_url,
    scope: scope
  });

  authenticator
    .authenticate(OfficeHelpers.DefaultEndpoints.Microsoft)
    .then(function (token) {
      access_token = token.access_token;
      //API呼び出し
      $(function () {
        $.ajax({
          url:
            "https://graph.microsoft.com/v1.0/sites/20531fc2-c6ab-4e1e-a532-9c8e15afed0d/drive/items/01SG44IHMJY6HM4OB2XJGZ34EYB77ZANB2",
          type: "GET",
          beforeSend: function (xhr) {
            xhr.setRequestHeader("Authorization", "Bearer " + access_token);
          }
        }).then(
          async function (data) {
            const obj = data["@microsoft.graph.downloadUrl"];
            var kakuinbase64 = await getImageBase64(obj);

            //ここからkakuinbase64を張り付ける処理
            inkanpaste(kakuinbase64);

            //ログ出力
            $(function () {
              Office.context.document.getFilePropertiesAsync(async function (asyncResult) {
                var fileUrl = asyncResult.value.url;
                var fileName;
                var inkanName;
                if (fileUrl == "") {
                  fileName = '未保存';
                  inkanName = "角印";
                } else {
                  fileName = fileUrl;
                  inkanName = "角印";
                };
                inkanLog(inkanName, fileName);
              });
            });

          },
          function (data) {
            console.log(data);
          }
        );
      });
    })
    .catch(OfficeHelpers.Utilities.log);
}

// バイナリ画像をbase64で返す
async function getImageBase64(url) {
  const response = await fetch(url);
  const contentType = response.headers.get("content-type");
  const arrayBuffer = await response.arrayBuffer();
  let base64String = btoa(String.fromCharCode.apply(null, new Uint8Array(arrayBuffer)));
  //return `data:${contentType};base64,${base64String}`;
  return base64String;
}

async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.error(error);
  }
}

Office.initialize = function (reason) {
  if (OfficeHelpers.Authenticator.isAuthDialog()) return;
};

async function inkanpaste(pic) {
  await Word.run(async (context) => {
    context.document.getSelection().insertInlinePictureFromBase64(pic, "End");
    await context.sync();
  });
}


//SharePointListにログ出力
function inkanLog(inkanName, inkanFile) {
  var authenticator;
  var client_id = "2b588edd-4adc-4e4f-a3e6-a8d362246a29";
  var redirect_url = "https://mikiyks.github.io/inkan4/";
  var scope = "https://graph.microsoft.com/Sites.ReadWrite.All";
  var access_token;

  authenticator = new OfficeHelpers.Authenticator();

  //access_token取得
  authenticator.endpoints.registerMicrosoftAuth(client_id, {
    redirectUrl: redirect_url,
    scope: scope
  });

  authenticator.authenticate(OfficeHelpers.DefaultEndpoints.Microsoft).then(function (token) {
    access_token = token.access_token;

    $(function () {
      $.ajax({
        url:
          "https://graph.microsoft.com/v1.0/sites/20531fc2-c6ab-4e1e-a532-9c8e15afed0d/lists/6aac0560-622e-4ee1-ba8f-73b32d8e9f05/items",
        type: "POST",
        data: JSON.stringify({
          fields: {
            Title: inkanName,
            FileName: inkanFile
          }
        }),
        contentType: "application/json",
        beforeSend: function (xhr) {
          xhr.setRequestHeader("Authorization", "Bearer " + access_token);
        }
      }).then(
        async function (data) { },
        function (data) {
          console.log(data);
        }
      );
    });
  });
}

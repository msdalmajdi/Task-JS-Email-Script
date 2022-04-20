// Your code goes here
// send email
//import { setApiKey, send } from "@sendgrid/mail";

"use strict";
const excelToJson = require("convert-excel-to-json");

const result = excelToJson({
  sourceFile: "names.xlsx",
});

const sgMail = require("@sendgrid/mail");
sgMail.setApiKey("");

//console.log(result["Sheet1"][0]["A"]);

result["Sheet1"].forEach((element) => {
  sendMessageWW(createMessageWW(element));
});

function createMessageWW(user) {
  const msg = {
    to: user["B"], // Change to your recipient
    from: "ruptural@hotmail.com", // Change to your verified sender
    subject: `Sending to  ${user["A"]}`,
    text: "and easy to do anywhere, even with Node.js",
    html: getCertificate(user),
  };
  return msg;
}
//`<strong>and easy to do anywhere, even with Node.js </strong> ${user["C"]}`
function sendMessageWW(msg) {
  sgMail
    .send(msg)
    .then(() => {
      console.log("Email sent");
    })
    .catch((error) => {
      console.error(error);
    });
}

function getCertificate(user) {
  var today = new Date();
  var date =
    today.getDate() + "/" + (today.getMonth() + 1) + "/" + today.getFullYear();
  return `<!DOCTYPE html>
    <html>
    <head>
    <title>Certificate</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@4.0.0/dist/css/bootstrap.min.css" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous">
    <link rel="stylesheet" href="styles.css">
    </head>
    <body>
        <div class="container-fluid p-4">
            <div class="row" style="height:100px">
                <div class="col-md-12">
                    <div class="row">
                        <div class="col-md-10">
                        </div>
                        <div class="col-md-2">
                            <img src="./assets/Microsoft_Azure_Logo.svg.png" style="width:100px;height:auto">
                        </div>
                    </div>
                </div>
            </div>
            <div class="row py-4">
                <div class="col-md-12">
                    <h3 class="text-center text-primary">
                        Certificate of Achievement
                    </h3>
                    <p class="text-center text-primary">
                        For getting an amazing grade in this ${user["C"]} course
                    </p>
                    <p class="text-center text-primary">
                       and Obtaining a grade of ${user["D"]}
                    </p>
                    <p class="text-center text-primary">we are happy to give this certificate to</p>
                    <h4 class="text-center text-primary">
                        ${user["A"]}
                    </h3>
    
                </div>
            </div>
            <div class="row">
                <div class="col-md-12">
                    <div class="row pt-4">
                        <div class="col-md-4">
                            <h3 class="text-left">
                                Bill gates
                            </h3>
                            <p class="text-left">
                                former ceo
                            </p>
                        </div>
                        <div class="col-md-4">
                        </div>
                        <div class="col-md-4">
                            <h3 class="text-right">
                                Jeff dean
                            </h3>
                            <p id='dDate'class="text-right">
                                ${date}
                            </p>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </body>
    </html>`;
}

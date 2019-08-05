/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

const setting1Name: string = 'setting1';
const setting2Name: string = 'setting2';
const setting3Name: string = 'setting3';

Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("setting1ValueSetter").onclick = setValue1;
    document.getElementById("setting2ValueSetter").onclick = setValue2;
    document.getElementById("setting3ValueSetter").onclick = setValue3;

    document.getElementById("setting1value").innerHTML = Office.context.roamingSettings.get(setting1Name);
    document.getElementById("setting2value").innerHTML = Office.context.roamingSettings.get(setting2Name);
    document.getElementById("setting3value").innerHTML = Office.context.roamingSettings.get(setting3Name);
  }
});

export async function setValue1() {
  let stringVal = document.getElementById("setting1value").innerHTML;
  let count = 1;
  
  if (stringVal !== undefined && stringVal !== "undefined" && stringVal !== "NaN") { 
    count = +stringVal + 1;
  }

  Office.context.roamingSettings.set(setting1Name, count);
  Office.context.roamingSettings.saveAsync();
  document.getElementById("setting1value").innerText = count.toString();
}

export async function setValue2() {
  let stringVal = document.getElementById("setting2value").innerHTML;
  let count = 1;
  
  if (stringVal !== undefined && stringVal !== "undefined" && stringVal !== "NaN") { 
    count = +stringVal + 1;
  }

  Office.context.roamingSettings.set(setting2Name, count);
  Office.context.roamingSettings.saveAsync();
  document.getElementById("setting2value").innerText = count.toString();
}

export async function setValue3() {
  let stringVal = document.getElementById("setting3value").innerHTML;
  let count = 1;
  
  if (stringVal !== undefined && stringVal !== "undefined" && stringVal !== "NaN") { 
    count = +stringVal + 1;
  }
  
  Office.context.roamingSettings.set(setting3Name, count);
  Office.context.roamingSettings.saveAsync();
  document.getElementById("setting3value").innerText = count.toString();
}

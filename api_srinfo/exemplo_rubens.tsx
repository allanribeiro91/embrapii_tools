const SHOW_LOG = true;
const log = (info) => SHOW_LOG ? Logger.log(info) : null;

const SRINFO = {
  LOGIN: {
    "username": "CONTA",
    "password": "SENHA",
  },
  API: {
    "Unidades": {
      urlApi: "https://srinfo.embrapii.org.br/units/api/list/",
      sheetName: "units",
    },
    "Projetos": {
      urlApi: "https://srinfo.embrapii.org.br/projects/api/projects/",
      sheetName: "projects",
    },
    "Contratos": {
      urlApi: "https://srinfo.embrapii.org.br/projects/api/contracts/",
      sheetName: "contracts",
    },
    "Empresas": {
      urlApi: "https://srinfo.embrapii.org.br/company/api/companies/",
      sheetName: "companies"
    },
    "Negociações": {
      urlApi: "https://srinfo.embrapii.org.br/units/api/negotiations/",
      sheetName: "negotiations",
    },
    "Estudantes": {
      urlApi: "https://srinfo.embrapii.org.br/people/api/students/",
      sheetName: "students",
    },
    "Pedidos de PI": {
      urlApi: "https://srinfo.embrapii.org.br/projectmonitoring/api/iprequests/",
      sheetName: "iprequests",
    },
    "Equipe": {
      urlApi: "https://srinfo.embrapii.org.br/people/api/roles/",
      sheetName: "roles",
    },
    "Prospecções": {
      urlApi: "https://srinfo.embrapii.org.br/units/api/prospect/",
      sheetName: "prospect",
    },
    "Termos de Cooperação": {
      urlApi: "https://srinfo.embrapii.org.br/accreditation/api/cooperationterm/",
      sheetName: "cooperationterm",
    },
    "Propostas técnicas": {
      urlApi: "https://srinfo.embrapii.org.br/units/api/negotiations/technicalproposals/",
      sheetName: "technicalproposals",
    }
  },
};

function getData() {
  const username = SRINFO.LOGIN.username;
  const password = SRINFO.LOGIN.password;
  let tokenApi = getToken(username, password);

  let page;
  for (const [itemApi, dataApi] of Object.entries(SRINFO.API)) {
    log("=======================================");
    log(`Importação de dados: ${itemApi}`);
    let data = [];
    let results = [];
    let url = dataApi.urlApi;
    page = 1;
    try {
      while (url != null) {
        log(`Página ${page}`);
        let json = getJSON(url, tokenApi);
        results = json.results;
        url = json.next;
        data = data.concat(results);
        sleep(1);
        page += 1;
      }
      insertData(SpreadsheetApp.getActiveSpreadsheet(), dataApi.sheetName, data);
      log(`Dados inseridos na planilha: ${itemApi}`);
      sleep(3);
      tokenApi = getToken(username, password);
    } catch (e) {
      log(e);
    }
  }
  updateTime();
}

function getToken(username, password) {
  const urlApiToken = "https://srinfo.embrapii.org.br/token/";
  let config = {
    headers: {},
    method: "post",
    payload: { username, password },
    muteHttpExceptions: false,
  };

  let response = UrlFetchApp.fetch(urlApiToken, config);
  let dataAll = JSON.parse(response.getContentText());

  return dataAll.refresh == null ? dataAll.detail : dataAll.access;
}

function getJSON(url, tokenApi) {
  let config = {
    headers: { "Authorization": `Bearer ${tokenApi}` },
    method: "get",
    muteHttpExceptions: false,
    contentType: "application/json",
  };

  let response = UrlFetchApp.fetch(url, config);
  let dataAll = JSON.parse(response.getContentText());

  for (let [rowId, row] of Object.entries(dataAll.results)) {
    for (let [entryId, entry] of Object.entries(row)) {
      if (Array.isArray(entry)) {
        dataAll["results"][rowId][entryId] = stringifyList(entry);
      }

      let monetaryValues = ['embrapii_amount', 'company_amount', 'ue_amount', 'total_amount']
      if (monetaryValues.includes(entryId)) {
        dataAll["results"][rowId][entryId] = entry.replace(".", ",");
      }
    }

    if (url.includes("companies") && !row.hasOwnProperty("primary_activity")) {
      row["primary_activity"] = null;
      dataAll["results"][rowId] = row;
    }
  }

  return dataAll;
}

function clearPageData(spreadsheet, sheetName) {
  spreadsheet.getSheetByName(sheetName).clearContents();
}

function insertData(spreadsheet, sheetName, json) {
  clearPageData(spreadsheet, sheetName);
  
  let sheet = spreadsheet.getSheetByName(sheetName);

  let header = Object.keys(json[0]);
  sheet.getRange("1:1").setValues([header]);

  let dataAll = [];
  for (const [key, value] of Object.entries(json)) {
    dataAll.push(header.map(h => value[h]));
  }

  let dataRange = "2:" + (dataAll.length + 1);
  sheet.getRange(dataRange).setValues(dataAll);
}

function stringifyList(list) {
  return list.join(";\n").trim();
}

function sleep(s) {
  log(`Pausa: ${s} segundos`);
  let date = Date.now();
  let currentDate = null;
  do {
    currentDate = Date.now();
  } while (currentDate - date < (s * 1000));
}

function updateTime() {
  let day = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy")
  let hour = Utilities.formatDate(new Date(), "GMT-3", "hh:mm")
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard").getRange("B1:B2").setValues([[day], [hour]]);
}

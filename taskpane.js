/* global Office  v.001 */

eamDataConverters = {
  "MIXVARCHAR": (val) => val,
  "VARCHAR": (val) => val,
  "NUMBER" : (val) => parseFloat(val),
  "LONG" : (val) => parseInt(val),
  "CONTAINS": (val) => val.split(','),
  "CURRENCY": (val) => parseFloat(val),
  "DECIMAL" : (val) => parseFloat(val),
  "DURATION" : (val) => parseInt(val), // Assuming duration in minutes or similar
  "DATE" : (val) => new Date(val),
  "DATETIME" : (val) => new Date(val),
  "TSMIDNIGHT" : (val) => new Date(val), // Assuming it's a timestamp at midnight
  "DEPENDENT" : (val) => val, // No specific conversion, just return the value
  "CLOB" : (val) => val, // Assuming CLOB is returned as a string
  "CHKBOOLEAN": (val) => val === '1' || val === 'true' || val === true
}
convert = (stringValue, dataType) => {
  const converter = eamDataConverters[dataType] || ((val) => val); // Default to identity function if no converter found
  return converter(stringValue);
};


loadDataFromEAM = async (
  tenant,//: string,
  organization,//: string,
  apiKey,//: string,
  gridName,//: string,
  numberOfRecord,//: number,
  parameters,//: [{name: string, type: string, value: object}]
) => {
  const eamData = await fetch( 'https://eu1.eam.hxgnsmartcloud.com/axis/restservices/grids', {
    method: 'POST',
    credentials: 'include',
    headers: {
      'Content-Type': 'application/json',
      'tenant': tenant,
      'organization': organization,
      'X-API-KEY': apiKey
    },
    body: JSON.stringify(
      {
        "GRID": { "GRID_NAME": gridName, "USER_FUNCTION_NAME": gridName, "NUMBER_OF_ROWS_FIRST_RETURNED": numberOfRecord || 1000, "CURSOR_POSITION": 0 },
        "GRID_TYPE": { "TYPE": "LIST" },
        "REQUEST_TYPE": "LIST.HEAD_DATA.STORED",

        "LOV": {
          "LOV_PARAMETERS": {
            "LOV_PARAMETER": (parameters || []).map( param => { return {
              "ALIAS_NAME": param.name,
              "TYPE": param.type,
              "VALUE": "" + param.value
            };})
          }
        }
      })
  });

  if ( eamData.ok ) {
      const eamJson = await eamData.json();
      const activityTypesOfWorkspaces = eamJson.Result?.ResultData?.DATARECORD?.map( record => {
        const result = {};
        // building first record object
        record.DATAFIELD.forEach( field => {
          result[field.FIELDNAME] = convert(field.FIELDVALUE, field.DATATYPE);
        });
        return result;
      });
      return { ok: true, data: activityTypesOfWorkspaces};
  }
  return { ok: false, detail: eamData};
}

Office.onReady(async (info) => {

  let tenant = "HXGNDEMO0016_DEM";
  let organization = "*";
  let apiKey = 'aefa7f2bf2-78d4-4d87-aac7-f2f6d2869f4a';

  const workspaces = {
    /*
    workspace_1: { 
      description : "",
      activityTypes: [
        { id: "", description: "", color: ""}
      ],
      engagementTypes: [
        ]
    }
     */
  };

  if ( info.host === Office.HostType.Outlook) {
    document.getElementById('btnSync').onclick = syncToTimeline;
    document.getElementById('btnCancel').onclick = closePane;
    
    const workspace = document.getElementById('workspace');
    const activityType = document.getElementById('activityType');
    const engagementType = document.getElementById('engagementType');
    const customerEvent = document.getElementById('customerEvent');
    
    const userEmailAddress = Office.context.mailbox.userProfile.emailAddress;
    

    let eamData = await loadDataFromEAM(tenant, organization, apiKey, "1UTLAC", 1000, [{ name: "param.user_email", type:"string", value: userEmailAddress}]);
    if ( eamData.ok ) {
      eamData.data.reduce( (acc, curr) => {
        // merging all records from the same workspace together
        acc[curr.c_workspace] = acc[curr.c_workspace] || { description: curr.c_description || "", default: curr.c_default || false, activityTypes: [], engagementTypes: [] };
        acc[curr.c_workspace].activityTypes.push({ 
          id: curr.type_id, 
          description: curr.type_desc,
          color: curr.type_color,
          gobal: curr.type_global

        });
        return acc;
      }, workspaces);
    } else {
      alert("Error fetching data from EAM: " + eamData.detail.status + " - " + eamData.detail.statusText);
      return;
    }

    eamData = await loadDataFromEAM(tenant, organization, apiKey, "1UTLEG", 1000, [{ name: "param.user_email", type:"string", value: userEmailAddress}]);
    if ( eamData.ok ) {
      eamData.data.reduce( (acc, curr) => {
        acc[curr.c_workspace] = acc[curr.c_workspace] || { description: curr.c_description || "", default: curr.c_default || false, activityTypes: [], engagementTypes: [] };
        acc[curr.c_workspace].engagementTypes.push({ 
          id: curr.type_id, 
          description: curr.type_desc,
          color: curr.type_color,
          gobal: curr.type_global});
        return acc;
      }, workspaces);
    } else {
      alert("Error fetching data from EAM: " + eamData.detail.status + " - " + eamData.detail.statusText);
      return;
    }


    workspace.addEventListener('change', (e) => {
      // we must empty all drop down menu
      activityType.innerHTML = '';
      engagementType.innerHTML = '';

      activityType.appendChild(new Option('-- Select Activity Type --', ''));
      engagementType.appendChild(new Option('-- Select Engagement Type --', ''));

      workspaces[defaultWorkspace].activityTypes.forEach( at => {
        const option = document.createElement('option');
        option.value = at.id;
        option.textContent = at.description;
        activityType.appendChild(option);
      });

      workspaces[defaultWorkspace].engagementTypes.forEach( et => {
        const option = document.createElement('option');
        option.value = et.id;
        option.textContent = et.description;
        engagementType.appendChild(option);
      });
      
    });

    workspace.innerHTML = '';
    Object.keys(workspaces).forEach( wsKey => {
      const ws = workspaces[wsKey];
      const option = document.createElement('option');
      option.value = wsKey;
      option.textContent = ws.description || wsKey;
      if ( ws.default ) {
        option.setAttribute('default', '');
        option.setAttribute('selected', 'selected');
      }
      workspace.appendChild(option);
    });

    //const defaultWorkspace = Object.keys(workspaces).filter( wsKey => workspaces[wsKey].default );

    activityType.addEventListener('change', validateForm);
    engagementType.addEventListener('change',  validateForm);
    customerEvent.addEventListener('input', validateForm);
    
	  // Initial load
	  loadExistingValues();
    
    // Run validation
	  validateForm();
  }
});

function validateForm() {
  const activityType = document.getElementById('activityType').value;
  const engagementType = document.getElementById('engagementType').value;
  const customerEvent = document.getElementById('customerEvent').value;
  const btnSync = document.getElementById('btnSync');
  
  // Disable engagement type if PTO is selected
  if (activityType === 'PTO') {
    document.getElementById('engagementType').value = '';
    document.getElementById('engagementType').disabled = true;
  } else {
    document.getElementById('engagementType').disabled = false;
  }
  
  // Enable sync button logic
  let isValid = false;
  if (activityType && customerEvent) {
    if (activityType === 'PTO') {
      isValid = true;
    } else if (engagementType) {
      isValid = true;
    }
  }
  
  btnSync.disabled = !isValid;
}

function loadExistingValues() {
  Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const customProps = result.value;
      
      const activityType = customProps.get('ActivityType');
      const engagementType = customProps.get('EngagementType');
      const customerEvent = customProps.get('CustomerEvent');
      const OnSite = customProps.get('OnSite');
      const CustInteraction = customProps.get('CustInteraction');
      const Clevel = customProps.get('Clevel');
      const FullDay = customProps.get('FullDay');
      
      if (activityType) document.getElementById('activityType').value = activityType;
      if (engagementType) document.getElementById('engagementType').value = engagementType;
      if (customerEvent) document.getElementById('customerEvent').value = customerEvent;
      if (OnSite === true || OnSite === 'true') document.getElementById('OnSite').checked = true;
      if (CustInteraction === true || CustInteraction === 'true') document.getElementById('CustInteraction').checked = true;
      if (Clevel === true || Clevel === 'true') document.getElementById('Clevel').checked = true;
      if (FullDay === true || FullDay === 'true') document.getElementById('FullDay').checked = true;
      
      // CRITICAL: Call validation AFTER the values are set
      validateForm();
    }
  });
}

function saveCustomProperties(callback) {
  Office.context.mailbox.item.loadCustomPropertiesAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const customProps = result.value;
      
      customProps.set('ActivityType', document.getElementById('activityType').value);
      customProps.set('EngagementType', document.getElementById('engagementType').value);
      customProps.set('CustomerEvent', document.getElementById('customerEvent').value);
      customProps.set('OnSite', document.getElementById('OnSite').checked);
      customProps.set('CustInteraction', document.getElementById('CustInteraction').checked);
      customProps.set('Clevel', document.getElementById('Clevel').checked);
      customProps.set('FullDay', document.getElementById('FullDay').checked);
      
      customProps.saveAsync((saveResult) => {
        if (saveResult.status === Office.AsyncResultStatus.Succeeded) {
          callback(true);
        } else {
          callback(false);
        }
      });
    }
  });
}

async function syncToTimeline() {
  const statusDiv = document.getElementById('status');
  statusDiv.className = 'status-message';
  statusDiv.style.display = 'none';
  
  // Save properties first
  saveCustomProperties(async (success) => {
    if (!success) {
      showStatus('Failed to save properties', 'error');
      return;
    }
    
    try {
      const item = Office.context.mailbox.item;
      // Get appointment data
      const appointmentData = await getAppointmentData(item);
      console.log("data.organizer:", appointmentData.organizer);
      
      // Build JSON payload
      const json = await buildJsonPayload(appointmentData);
      
      // Send to API
      const response = await fetch('https://dataflow-inbound-message-prd-euc1.eam.hxgnsmartcloud.com/api/message?tag=timeline', {
        method: 'POST',
        headers: {
          'accept': 'application/json',
          'X-Tenant-Id': 'HXGNDEMO0016_DEM',
          'Authorization': 'Basic SDNBV0JNX0hYR05ERU1PMDAxNl9ERU06RyFvYmEhMjAyMA==',
          'Content-Type': 'text/plain'
        },
        body: json
      });
      
      if (response.ok) {
        // showStatus('Appointment sent to Timeline successfully!\nClick on "Open Timeline Tenant" or "Close"', 'success');
        const msgText = await response.text();
		showStatus(msgText, 'success'); 
      } else {
        const errorText = await response.text();
		const msgTextErr1 = `Error: ${response.status} - ${errorText}`;
        showStatus(msgTextErr1, 'error');
      }
    } catch (error) {
		const msgTextErr2 = `Error: ${error.message}`;
        showStatus(msgTextErr2, 'error');
    }
  });
}

async function getAppointmentData(item) {
  return new Promise(async (resolve) => {
    // Helper to handle the "Compose vs Read" difference for strings/dates
    const getValue = async (property) => {
      if (property && typeof property.getAsync === 'function') {
        return new Promise(r => property.getAsync(result => r(result.value || '')));
      }
      return property || '';
    };
	  
    const data = {
      subject: await getValue(item.subject),
      location: await getValue(item.location),
      start: await getValue(item.start),
      end: await getValue(item.end),
      organizer: '',
      body: ''
    };

    // 1. Handle Organizer (It's a bit more complex)
    if (item.organizer) {
      if (typeof item.organizer.getAsync === 'function') {
        const orgRes = await new Promise(r => item.organizer.getAsync(r));
        data.organizer = orgRes.value ? (orgRes.value.emailAddress || orgRes.value.displayName) : '';
      } else {
        data.organizer = item.organizer.emailAddress || item.organizer.displayName || '';
      }
    }

    // 2. Get Body (Always Async)
    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        data.body = result.value || '';
      }

      // 3. Get Custom Properties
      item.loadCustomPropertiesAsync((propResult) => {
        if (propResult.status === Office.AsyncResultStatus.Succeeded) {
          const props = propResult.value;
          // Keeping keys short (8-char logic)
          data.actType = props.get('ActivityType') || '';
          data.engType = props.get('EngagementType') || '';
          data.custEvt = props.get('CustomerEvent') || data.subject;
		  data.OnSite = props.get('OnSite');
		  data.CustInteraction = props.get('CustInteraction');
		  data.Clevel = props.get('Clevel');
		  data.FullDay = props.get('FullDay');
        }
        resolve(data);
      });
    });
  });
}

async function buildJsonPayload(data) {
  // Get owner email (use current user as fallback)
  const ownerEmail = Office.context.mailbox.userProfile.emailAddress;
  
  // Parse organizer info
  let aliasStr = 'no alias';
  let firstNameStr = 'External';
  let lastNameStr = 'External';
  
  if (data.organizer) {
    if (data.organizer.includes('@')) {
      aliasStr = data.organizer;
      lastNameStr = data.organizer.split('@')[0];
      firstNameStr = '';
    } else {
      const parts = data.organizer.split(' ');
      if (parts.length >= 2) {
        firstNameStr = parts[0];
        lastNameStr = parts.slice(1).join(' ');
      } else {
        lastNameStr = data.organizer;
      }
    }
  }
  
  // Format dates
  const formatDate = (date) => {
    const d = new Date(date);
    const day = String(d.getDate()).padStart(2, '0');
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const year = d.getFullYear();
    const hours = String(d.getHours()).padStart(2, '0');
    const minutes = String(d.getMinutes()).padStart(2, '0');
    const seconds = String(d.getSeconds()).padStart(2, '0');
	const myDate = `${day}/${month}/${year} ${hours}:${minutes}:${seconds}`;
    return myDate;
  };
  
  // Clean body
  let cleanBody = data.body.replace(/[\r\n]+/g, ' ').replace(/"/g, ' ').replace(/[{}\[\]]/g, ' ').substring(0, 255); 
  if (data.FullDay) {
	  cleanBody = "Orig. " + data.start + " -> " + data.end + " - " +  cleanBody;
	  cleanBody = cleanBody.substring(0, 250);
  }
// Handle PTO
  let customerEvent = data.custEvt; // Start with the captured subject
  let engagementType = data.engType;
  if (data.actType === 'PTO') {
    customerEvent = 'Personal Time OFF'; // Use the local variable
    engagementType = '';
  }
  const Location =  data.location || '';
  const CreationTime = new Date().toISOString();
  const OnSite = (data.OnSite ?? false).toString();
  const CustInteraction = (data.CustInteraction ?? false).toString();
  const Clevel = (data.Clevel ?? false).toString();
  const FullDay = (data.FullDay ?? false).toString();

  // EntryID (Standard Office.js itemId)
  // Logic fix for IDs:
  const entryID = Office.context.mailbox.item.itemId || await new Promise(r => Office.context.mailbox.item.saveAsync(res => r(res.value)));
  
  // 1. Initialize globalID with entryID as the safe fallback
  let globalID = entryID;
  
  // 2. Only attempt to fetch headers if the function exists (Read Mode)
  if (Office.context.mailbox.item.getAllInternetHeadersAsync) {
      globalID = await new Promise((resolve) => {
          Office.context.mailbox.item.getAllInternetHeadersAsync((result) => {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                  const headers = result.value || {};
                  // Resolve with UID if found, otherwise the entryID
                  resolve(headers["UID"] || headers["vcal-uid"] || entryID);
              } else {
                  resolve(entryID);
              }
          });
      });
  } else {
      // 3. In Compose Mode, headers don't exist yet. 
      // The entryID generated by your saveAsync() call is the best available unique ID.
      console.log("Internet headers not available in Compose mode; using entryID.");
      globalID = entryID;
  }

  const payload = {
    EntryID: entryID,
    globalID: globalID,
    Organizer: data.organizer,
    AuthorAlias: aliasStr,
    AuthorFirstname: firstNameStr,
    AuthorLastname: lastNameStr,
    OwnerEmail: ownerEmail,
    Subject: customerEvent,
    Start: formatDate(data.start),
    End: formatDate(data.end),
    Location: Location,
    CreationTime: CreationTime,
    ActivityType: data.actType,
    EngagementType: engagementType,
    OnSite: OnSite,
    CustInteraction: CustInteraction,
    Clevel: Clevel,
	FullDay: FullDay,
    Note: cleanBody
  };
  
  return JSON.stringify(payload);
}

function showStatus(message, type) {
  const statusDiv = document.getElementById('status');
  statusDiv.textContent = message;
  statusDiv.className = `status-message ${type}`;
  statusDiv.style.display = 'block';
}

function closePane() {
  Office.context.ui.closeContainer();


}


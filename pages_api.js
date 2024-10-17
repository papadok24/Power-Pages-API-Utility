(function (global) {
  const powerPagesWebApi = (global.powerPagesWebApi =
    global.powerPagesWebApi || {});

  // Function to get common headers with cached CSRF token
  const memoizedGetCachedCommonHeaders = (() => {
    let cachedHeaders = null;
    return async () => {
      if (!cachedHeaders) {
        const token = await shell.getTokenDeferred();
        cachedHeaders = {
          __RequestVerificationToken: token,
          Accept: 'application/json',
          Prefer: 'odata.include-annotations=*',
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0',
        };
      }
      return cachedHeaders;
    };
  })();

  // Function to safely make fetch requests with CSRF token
  async function safePowerPagesFetch(url, options = {}) {
    try {
      // Get common headers for request validation
      options.headers = {
        ...options.headers,
        ...(await memoizedGetCachedCommonHeaders()),
      };
      options.method = options.method || 'GET';
      options.credentials = 'include';

      // Retry mechanism for transient network errors
      const maxRetries = 3;
      for (let attempt = 1; attempt <= maxRetries; attempt++) {
        try {
          // Make the fetch request
          const response = await fetch(url, options);

          // Handle response based on status
          if (response.ok) {
            const jsonResponse =
              response.status === 204 ? null : await response.json();
            return simplifyGetResponse(jsonResponse);
          } else {
            await handleErrorResponse(response);
          }
        } catch (error) {
          if (attempt < maxRetries) {
            console.warn(`Attempt ${attempt} failed. Retrying...`);
          } else {
            console.error('Request failed after multiple attempts:', error);
            throw error;
          }
        }
      }
    } catch (error) {
      console.error('Request failed:', error);
      throw error;
    }
  }

  // Function to handle HTTP errors based on response status
  const errorMessages = new Map([
    [
      400,
      'Bad Request: The request could not be understood or was missing required parameters.',
    ],
    [401, 'Unauthorized: Missing or invalid authentication credentials.'],
    [403, 'Forbidden: You do not have permission to perform this action.'],
    [404, 'Not Found: The requested resource could not be found.'],
    [
      413,
      'Payload Too Large: The request entity is larger than the server is able to process.',
    ],
    [500, 'Internal Server Error: An unexpected error occurred on the server.'],
    [
      503,
      'Service Unavailable: The server is currently unable to handle the request.',
    ],
  ]);

  class PowerPagesApiError extends Error {
    constructor(message, status, details) {
      super(message);
      this.name = 'PowerPagesApiError';
      this.status = status;
      this.details = details;
    }
  }

  async function handleErrorResponse(response) {
    const errorMessage = errorMessages.get(response.status) || `Unexpected error occurred: ${response.status} - ${response.statusText}`;
    let details;
    try {
      details = await response.json();
    } catch (e) {
      details = null;
    }
    throw new PowerPagesApiError(errorMessage, response.status, details);
  }

  // Function to create a new record in the specified entity set
  async function createRecord(entitySetName, data) {
    const url = `/_api/${entitySetName}`;
    const options = {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(data),
    };
    return await safePowerPagesFetch(url, options);
  }

  // Function to update a record by ID in the specified entity set
  async function updateRecord(entitySetName, id, data) {
    const url = `/_api/${entitySetName}(${id})`;
    const options = {
      method: 'PATCH',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(data),
    };
    return await safePowerPagesFetch(url, options);
  }

  // Function to delete a record by ID in the specified entity set
  async function deleteRecord(entitySetName, id) {
    const url = `/_api/${entitySetName}(${id})`;
    const options = {
      method: 'DELETE',
    };
    return await safePowerPagesFetch(url, options);
  }

  // Function to get a specific record by ID with optional selected fields
  async function getRecord(entitySetName, id, select = null) {
    let url = `/_api/${entitySetName}(${id})`;
    if (select) {
      url += `?$select=${select}`;
    }
    return await safePowerPagesFetch(url);
  }

  // Function to query records with various query options (e.g., select, filter)
  async function queryRecords(entitySetName, queryOptions = {}) {
    const url = new URL(`/_api/${entitySetName}`, window.location.origin);
    const params = new URLSearchParams();

    if (queryOptions.select)
      params.append('$select', queryOptions.select.join(','));
    if (queryOptions.filter) params.append('$filter', queryOptions.filter);
    if (queryOptions.top) params.append('$top', queryOptions.top);
    if (queryOptions.orderby) params.append('$orderby', queryOptions.orderby);
    if (queryOptions.expand) params.append('$expand', queryOptions.expand);
    if (queryOptions.count) params.append('$count', 'true');

    url.search = params.toString();
    return await safePowerPagesFetch(url.toString());
  }

  // Function to execute FetchXML queries
  async function queryRecordsWithFetchXml(entitySetName, fetchXml) {
    const encodedFetchXml = encodeURIComponent(fetchXml);
    const url = `/_api/${entitySetName}?fetchXml=${encodedFetchXml}`;
    return await safePowerPagesFetch(url);
  }

  // Function to associate two records via navigation property
  async function associateRecords(
    primaryEntitySetName,
    primaryId,
    navigationProperty,
    relatedEntitySetName,
    relatedId,
  ) {
    const url = `/_api/${primaryEntitySetName}(${primaryId})/${navigationProperty}/$ref`;
    const options = {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        '@odata.id': `/_api/${relatedEntitySetName}(${relatedId})`,
      }),
    };
    return await safePowerPagesFetch(url, options);
  }

  // Function to disassociate two records via navigation property
  async function disassociateRecords(
    primaryEntitySetName,
    primaryId,
    navigationProperty,
    relatedId,
  ) {
    const url = `/_api/${primaryEntitySetName}(${primaryId})/${navigationProperty}/$ref?$id=/_api/${relatedId}`;
    const options = {
      method: 'DELETE',
    };
    return await safePowerPagesFetch(url, options);
  }

  // Function to upload a file to Azure Blob Storage in chunks
  async function uploadFileToAzure(entityName, entityId, file) {
    const chunkSize = 50 * 1024 * 1024; // Define chunk size (50MB)
    const url = `/_api/file/InitializeUpload/${entityName}(${entityId})/blob`;
    let token;

    // Initialize file upload to get the token
    try {
      const initResponse = await safePowerPagesFetch(url, {
        method: 'POST',
        headers: {
          'x-ms-file-name': file.name,
          'x-ms-file-size': file.size,
        },
        body: '',
      });
      token = initResponse;
    } catch (error) {
      console.error('Failed to initialize upload:', error);
      throw error;
    }

    // Upload file in chunks concurrently
    const uploadPromises = [];
    for (
      let blockno = 0;
      blockno < Math.ceil(file.size / chunkSize);
      blockno++
    ) {
      const end = Math.min((blockno + 1) * chunkSize, file.size);
      const content = file.slice(blockno * chunkSize, end);

      uploadPromises.push(
        safePowerPagesFetch(
          `/_api/file/UploadBlock/blob?offset=${blockno * chunkSize}&fileSize=${
            file.size
          }&chunkSize=${chunkSize}&token=${token}`,
          {
            method: 'PUT',
            headers: {
              'x-ms-file-name': file.name,
            },
            body: content,
          },
        ).catch((error) => {
          console.error(`Failed to upload chunk ${blockno + 1}:`, error);
          throw error;
        }),
      );
    }

    try {
      const results = await Promise.allSettled(uploadPromises);
      const failedUploads = results.filter(
        (result) => result.status === 'rejected',
      );
      if (failedUploads.length > 0) {
        console.error(`${failedUploads.length} chunk(s) failed to upload`);
        throw new Error('Some chunks failed to upload');
      }
    } catch (error) {
      console.error('Failed to upload file in chunks:', error);
      throw error;
    }
  }

  // Function to simplify GET response data
  function simplifyGetResponse(response) {
    if (response && response.value) {
      return {
        formattedData: response.value.map((record) => {
          return Object.fromEntries(
            Object.entries(record).filter(
              ([key]) =>
                !key.startsWith('@odata.') &&
                !key.startsWith('@Microsoft.Dynamics.CRM.') &&
                key !== '@odata.etag',
            ),
          );
        }),
        metadata: {
          count: response['@Microsoft.Dynamics.CRM.totalrecordcount'],
          countLimitExceeded:
            response['@Microsoft.Dynamics.CRM.totalrecordcountlimitexceeded'],
        },
      };
    }
    return response;
  }

  // Function to update a single property value for a specified entity
  async function updateSingleProperty(entitySetName, id, propertyName, propertyValue) {
    const url = `/_api/${entitySetName}(${id})/${propertyName}`;
    const options = {
      method: 'PUT',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ value: propertyValue }),
    };
    return await safePowerPagesFetch(url, options);
  }

  // Function to delete or clear a field value for a specified entity
  async function deleteFieldValue(entitySetName, id, fieldName) {
    const url = `/_api/${entitySetName}(${id})/${fieldName}`;
    const options = {
      method: 'DELETE',
    };
    return await safePowerPagesFetch(url, options);
  }

  powerPagesWebApi.safePowerPagesFetch = safePowerPagesFetch;
  powerPagesWebApi.createRecord = createRecord;
  powerPagesWebApi.updateRecord = updateRecord;
  powerPagesWebApi.deleteRecord = deleteRecord;
  powerPagesWebApi.getRecord = getRecord;
  powerPagesWebApi.queryRecords = queryRecords;
  powerPagesWebApi.queryRecordsWithFetchXml = queryRecordsWithFetchXml;
  powerPagesWebApi.associateRecords = associateRecords;
  powerPagesWebApi.disassociateRecords = disassociateRecords;
  powerPagesWebApi.uploadFileToAzure = uploadFileToAzure;
  powerPagesWebApi.simplifyGetResponse = simplifyGetResponse;
  powerPagesWebApi.updateSingleProperty = updateSingleProperty;
  powerPagesWebApi.deleteFieldValue = deleteFieldValue;
})(window);

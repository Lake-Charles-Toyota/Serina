# Serina
Lake Charles Toyota Test AI

## Usage

This Azure Function exposes an HTTP endpoint for listing files in a SharePoint
site and retrieving their contents. To call the function:

```
GET /api/HttpTrigger1?list=true        # List available files
GET /api/HttpTrigger1?fileId=<id>      # Fetch and parse a file by ID
```

Optional query parameters:

- `summary=true` – return only the first 2000 characters of the file
- `debug=true` – include the raw Graph API URL in the response

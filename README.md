# Microsoft To Do exporter

This is simple Powershell script/application to export data from Microsoft To Do.

## Configuration

- rename `config.json.example` to `config.json`
- paste your application data (example [manifest.json](manifest.json))

#### Scopes

| Id                                   | Permission name |                       Description |
| :----------------------------------- | :-------------- | --------------------------------: |
| f45671fb-e0fe-4b4b-be20-3d3ce43f1bcb | Tasks.Read      | Allows the app to read your tasks |

## Export

- run `main.ps1`
- open browser using `http://localhost:8080/` and follow instructions
  - approve application to get data from your account
  - invoke export

Exported data will be exported to out folder.

# Microsoft 365 Scope

`UiPath.MicrosoftOffice365.Activities.Office365ApplicationScope`

Opens a connection to Microsoft 365 and authenticates all contained activities. This is the top-level scope activity required for non-connection-service activities.

**Connector:** `uipath-microsoft-365`

## Properties

### Configuration

| Name | Display Name | Type | Default | Description |
|------|-------------|------|---------|-------------|
| `ConnectionId` | Connection ID | `String` | | The connection identifier from UiPath Integration Service. |

### Body

Activities placed inside the scope are authenticated via the configured Microsoft 365 connection.

## XAML Example

```xml
<o365:Office365ApplicationScope DisplayName="Microsoft 365 Scope" ConnectionId="[connectionId]">
    <Sequence DisplayName="Do">
        <!-- Child activities go here -->
    </Sequence>
</o365:Office365ApplicationScope>
```

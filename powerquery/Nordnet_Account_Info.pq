let
    // =============================================================
    // Nordnet Account Info Query
    // Endpoint: GET /api/2/accounts/{accid}/info
    // =============================================================

    // Required parameters (create these in Manage Parameters)
    pAccId = Text.From(#"pAccId"),
    pBaseUrl = Text.From(#"pBaseUrl"),
    pApiVersionPath = Text.From(#"pApiVersionPath"),
    pTimeoutSeconds = Number.From(#"pTimeoutSeconds"),

    // Optional authentication/session header parameters
    pAuthHeaderName = Text.From(#"pAuthHeaderName"),
    pAuthHeaderValue = Text.From(#"pAuthHeaderValue"),
    pSessionHeaderName = Text.From(#"pSessionHeaderName"),
    pSessionHeaderValue = Text.From(#"pSessionHeaderValue"),

    // Build endpoint path and URL parts
    RelativePath = Text.Combine(
        {
            Text.TrimEnd(pApiVersionPath, "/"),
            "/accounts/",
            pAccId,
            "/info"
        },
        ""
    ),

    // Build headers with optional auth/session placeholders
    AuthHeaderRecord =
        if Text.Length(Text.Trim(pAuthHeaderName)) > 0 and Text.Length(Text.Trim(pAuthHeaderValue)) > 0
        then Record.AddField([], pAuthHeaderName, pAuthHeaderValue)
        else [],

    SessionHeaderRecord =
        if Text.Length(Text.Trim(pSessionHeaderName)) > 0 and Text.Length(Text.Trim(pSessionHeaderValue)) > 0
        then Record.AddField([], pSessionHeaderName, pSessionHeaderValue)
        else [],

    Headers = Record.Combine(
        {
            [Accept = "application/json"],
            AuthHeaderRecord,
            SessionHeaderRecord
        }
    ),

    // API call with controlled timeout and manual status handling
    RawResponse =
        Web.Contents(
            pBaseUrl,
            [
                RelativePath = RelativePath,
                Headers = Headers,
                Timeout = #duration(0, 0, 0, pTimeoutSeconds),
                ManualStatusHandling = {400, 401, 403, 404, 429, 500}
            ]
        ),

    ResponseStatus = try Value.Metadata(RawResponse)[Response.Status] otherwise 200,

    ParsedJson =
        if ResponseStatus >= 200 and ResponseStatus < 300 then
            Json.Document(RawResponse)
        else
            error Error.Record(
                "Nordnet API Error",
                "API kald fejlede",
                [Status = ResponseStatus, Endpoint = RelativePath]
            ),

    JsonRecord =
        if Value.Is(ParsedJson, type record)
        then ParsedJson
        else error Error.Record("Unexpected Payload", "Forventede JSON record.", [PayloadType = Value.Type(ParsedJson)]),

    // Flatten top-level fields into a key/value table
    AsTable = Record.ToTable(JsonRecord),
    RenamedColumns = Table.RenameColumns(AsTable, {{"Name", "Field"}, {"Value", "Value"}}),

    // Add refresh timestamp (UTC)
    WithRefreshTimestamp = Table.AddColumn(
        RenamedColumns,
        "LastRefreshUtc",
        each DateTimeZone.UtcNow(),
        type datetimezone
    ),

    // Optional: normalize scalar values to text for stable Excel table rendering
    ValueAsText = Table.TransformColumns(
        WithRefreshTimestamp,
        {
            {
                "Value",
                each
                    if Value.Is(_, type text) then _
                    else if Value.Is(_, type number) then Number.ToText(_)
                    else if Value.Is(_, type logical) then Logical.ToText(_)
                    else if Value.Is(_, type datetimezone) then DateTimeZone.ToText(_)
                    else if Value.Is(_, type datetime) then DateTime.ToText(_)
                    else if Value.Is(_, type date) then Date.ToText(_)
                    else if Value.Is(_, type time) then Time.ToText(_)
                    else Text.FromBinary(Json.FromValue(_), TextEncoding.Utf8),
                type text
            }
        }
    )
in
    ValueAsText

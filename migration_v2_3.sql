-- Migration v2.3 — App uptime tracking
-- Run once in SSMS against SNP1 database

-- Table: one row per production.exe launch
CREATE TABLE [dbo].[app_uptime] (
    [session_id]   INT           IDENTITY(1,1) PRIMARY KEY,
    [pc_ip]        NVARCHAR(15)  NOT NULL,
    [hostname]     NVARCHAR(100) NULL,
    [machine_name] NVARCHAR(50)  NULL,        -- set after operator selects line at login
    [app_version]  NVARCHAR(20)  NULL,
    [start_time]   DATETIME      NOT NULL DEFAULT GETDATE(),
    [end_time]     DATETIME      NULL          -- NULL = still running / crashed
);
GO

-- Called on app start; returns new session_id
CREATE PROCEDURE [dbo].[sp_AppStarted]
    @PcIp       NVARCHAR(15),
    @Hostname   NVARCHAR(100),
    @AppVersion NVARCHAR(20)
AS
BEGIN
    SET NOCOUNT ON;
    INSERT INTO app_uptime (pc_ip, hostname, app_version, start_time)
    VALUES (@PcIp, @Hostname, @AppVersion, GETDATE());
    SELECT SCOPE_IDENTITY() AS session_id;
END
GO

-- Called after operator selects Line 1 / Line 2 at login
CREATE PROCEDURE [dbo].[sp_AppSetMachine]
    @SessionId   INT,
    @MachineName NVARCHAR(50)
AS
BEGIN
    SET NOCOUNT ON;
    UPDATE app_uptime SET machine_name = @MachineName WHERE session_id = @SessionId;
END
GO

-- Called on clean app close
CREATE PROCEDURE [dbo].[sp_AppStopped]
    @SessionId INT
AS
BEGIN
    SET NOCOUNT ON;
    UPDATE app_uptime SET end_time = GETDATE() WHERE session_id = @SessionId;
END
GO

-- Useful query: view uptime history per line
-- SELECT
--     session_id,
--     ISNULL(machine_name, '(not set)') AS line,
--     hostname,
--     pc_ip,
--     app_version,
--     start_time,
--     ISNULL(end_time, GETDATE()) AS end_time,
--     CASE WHEN end_time IS NULL THEN 'Running' ELSE 'Stopped' END AS status,
--     CAST(DATEDIFF(second, start_time, ISNULL(end_time, GETDATE())) / 3600.0
--          AS DECIMAL(10,2)) AS uptime_hours
-- FROM app_uptime
-- ORDER BY start_time DESC;

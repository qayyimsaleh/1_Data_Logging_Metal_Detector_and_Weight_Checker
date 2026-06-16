-- ============================================================
-- migration_v2_1.sql
-- SNP1 Database Migration — v2.0 -> v2.1
-- Run once against the SNP1 database on SQLEXPRESS.
-- ============================================================
USE [SNP1];
GO

-- ============================================================
-- 1. production_session — add end_date, end_by_user columns
-- ============================================================
IF NOT EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('dbo.production_session')
      AND name = 'end_date'
)
BEGIN
    ALTER TABLE dbo.production_session ADD end_date DATETIME NULL;
    ALTER TABLE dbo.production_session ADD end_by_user NVARCHAR(10) NULL;
    PRINT 'production_session: end_date, end_by_user columns added.';
END
ELSE
    PRINT 'production_session: end_date already exists, skipped.';
GO

-- ============================================================
-- 2. production_session — add PRIMARY KEY on production_id
--    Check for duplicates first. If this fails, resolve
--    duplicates manually before re-running.
-- ============================================================

-- Informational: show any duplicates that would block the PK
SELECT production_id, COUNT(*) AS cnt
FROM dbo.production_session
GROUP BY production_id
HAVING COUNT(*) > 1;

IF NOT EXISTS (
    SELECT 1 FROM sys.key_constraints
    WHERE parent_object_id = OBJECT_ID('dbo.production_session')
      AND type = 'PK'
)
BEGIN
    ALTER TABLE dbo.production_session
        ADD CONSTRAINT PK_production_session PRIMARY KEY (production_id);
    PRINT 'production_session: PRIMARY KEY added.';
END
ELSE
    PRINT 'production_session: PRIMARY KEY already exists, skipped.';
GO

-- ============================================================
-- 3. SEQUENCE objects — race-safe ID generation
--    Replaces MAX(id)+1 pattern in sp_GetNextProductionId
--    and sp_GetNextLogId.
-- ============================================================

-- seq_production_id: seed from current max + 1
IF NOT EXISTS (SELECT 1 FROM sys.sequences WHERE name = 'seq_production_id')
BEGIN
    DECLARE @MaxProd INT;
    SELECT @MaxProd = ISNULL(MAX(production_id), 0) + 1 FROM dbo.production_session;
    EXEC('CREATE SEQUENCE dbo.seq_production_id
              START WITH ' + @MaxProd + '
              INCREMENT BY 1
              NO CYCLE
              NO CACHE;');
    PRINT 'Sequence seq_production_id created.';
END
ELSE
    PRINT 'seq_production_id already exists, skipped.';
GO

-- seq_log_id: seed from current max + 1
IF NOT EXISTS (SELECT 1 FROM sys.sequences WHERE name = 'seq_log_id')
BEGIN
    DECLARE @MaxLog INT;
    SELECT @MaxLog = ISNULL(MAX(log_id), 0) + 1 FROM dbo.main_logging;
    EXEC('CREATE SEQUENCE dbo.seq_log_id
              START WITH ' + @MaxLog + '
              INCREMENT BY 1
              NO CYCLE
              NO CACHE;');
    PRINT 'Sequence seq_log_id created.';
END
ELSE
    PRINT 'seq_log_id already exists, skipped.';
GO

-- ============================================================
-- 4. sp_GetNextProductionId — use SEQUENCE (atomic, no race)
-- ============================================================
ALTER PROCEDURE [dbo].[sp_GetNextProductionId]
AS
BEGIN
    SET NOCOUNT ON;
    SELECT NEXT VALUE FOR dbo.seq_production_id;
END
GO

-- ============================================================
-- 5. sp_GetNextLogId — use SEQUENCE (atomic, no race)
-- ============================================================
ALTER PROCEDURE [dbo].[sp_GetNextLogId]
AS
BEGIN
    SET NOCOUNT ON;
    SELECT NEXT VALUE FOR dbo.seq_log_id;
END
GO

-- ============================================================
-- 6. sp_UpdateMachineConfig — persist config from app to DB
-- ============================================================
IF OBJECT_ID('dbo.sp_UpdateMachineConfig', 'P') IS NOT NULL
    DROP PROCEDURE dbo.sp_UpdateMachineConfig;
GO

CREATE PROCEDURE [dbo].[sp_UpdateMachineConfig]
    @PcIp           NVARCHAR(15),
    @WeigherIp      NVARCHAR(15),
    @MetalDetectorIp NVARCHAR(50) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    UPDATE dbo.machine_location
    SET    weigher_ip          = @WeigherIp,
           metal_detector_ip   = @MetalDetectorIp
    WHERE  pc_ip = @PcIp;
END
GO

-- ============================================================
-- 7. sp_EndProductionSession — stamp end_date when session stops
-- ============================================================
IF OBJECT_ID('dbo.sp_EndProductionSession', 'P') IS NOT NULL
    DROP PROCEDURE dbo.sp_EndProductionSession;
GO

CREATE PROCEDURE [dbo].[sp_EndProductionSession]
    @ProductionId   INT,
    @EndDate        DATETIME,
    @EndByUser      NVARCHAR(10)
AS
BEGIN
    SET NOCOUNT ON;
    UPDATE dbo.production_session
    SET    end_date     = @EndDate,
           end_by_user  = @EndByUser
    WHERE  production_id = @ProductionId
      AND  end_date IS NULL;
END
GO

-- ============================================================
-- 8. Password hashing — SHA-256 hex (64 chars)
--
--    IMPORTANT: After running this block, ALL existing passwords
--    are invalidated because the column format changes from
--    plain text to SHA-256 hex. Admin must reset every user
--    password using sp_SetUserPassword before operators can
--    log in again.
--
--    To compute SHA-256 hex of a password in Python:
--        import hashlib
--        hashlib.sha256("yourpassword".encode()).hexdigest()
-- ============================================================

-- Expand column to hold 64-char hex digest
IF EXISTS (
    SELECT 1 FROM sys.columns
    WHERE object_id = OBJECT_ID('dbo.users')
      AND name = 'password'
      AND max_length < 128   -- nchar(50) = 100 bytes; nvarchar(64) = 128 bytes
)
BEGIN
    ALTER TABLE dbo.users ALTER COLUMN password NVARCHAR(64) NOT NULL;
    PRINT 'users.password column expanded to NVARCHAR(64).';
    PRINT '*** ACTION REQUIRED: Reset all user passwords via sp_SetUserPassword ***';
END
ELSE
    PRINT 'users.password already NVARCHAR(64)+, skipped.';
GO

-- ============================================================
-- 9. sp_VerifyUser — compare against SHA-256 hash
-- ============================================================
ALTER PROCEDURE [dbo].[sp_VerifyUser]
    @Username     NVARCHAR(50),
    @PasswordHash NVARCHAR(64)   -- SHA-256 hex from Python
AS
BEGIN
    SET NOCOUNT ON;
    SELECT username
    FROM   dbo.users
    WHERE  username = @Username
      AND  password = @PasswordHash;
END
GO

-- ============================================================
-- 10. sp_SetUserPassword — admin tool to set/reset a user password
--     Pass the SHA-256 hex hash of the new password.
--     Example (compute in Python first):
--         import hashlib
--         h = hashlib.sha256("Pass1234".encode()).hexdigest()
--     Then:
--         EXEC sp_SetUserPassword 'operator1', '<hex_from_python>'
-- ============================================================
IF OBJECT_ID('dbo.sp_SetUserPassword', 'P') IS NOT NULL
    DROP PROCEDURE dbo.sp_SetUserPassword;
GO

CREATE PROCEDURE [dbo].[sp_SetUserPassword]
    @Username     NVARCHAR(50),
    @PasswordHash NVARCHAR(64)
AS
BEGIN
    SET NOCOUNT ON;
    IF EXISTS (SELECT 1 FROM dbo.users WHERE username = @Username)
        UPDATE dbo.users SET password = @PasswordHash WHERE username = @Username;
    ELSE
        INSERT INTO dbo.users (username, password) VALUES (@Username, @PasswordHash);
END
GO

PRINT '=== migration_v2_1.sql complete ===';
GO

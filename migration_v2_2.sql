-- Migration v2.2 — Machine/Line management SPs
-- Run once in SSMS against SNP1 database

-- 1. Get machine config by name
CREATE PROCEDURE [dbo].[sp_GetMachineByName]
    @MachineName NVARCHAR(50)
AS
BEGIN
    SET NOCOUNT ON;
    SELECT machine, weigher_ip, metal_detector_ip
    FROM machine_location
    WHERE machine = @MachineName;
END
GO

-- 2. Get all machines with full detail (for Manage Lines screen)
CREATE PROCEDURE [dbo].[sp_GetAllMachinesDetail]
AS
BEGIN
    SET NOCOUNT ON;
    SELECT machine, pc_ip, weigher_ip, metal_detector_ip
    FROM machine_location
    ORDER BY machine;
END
GO

-- 3. Add new machine/line
CREATE PROCEDURE [dbo].[sp_AddMachine]
    @Machine        NVARCHAR(50),
    @PcIp           NVARCHAR(15),
    @WeigherIp      NVARCHAR(15),
    @MetalIp        NVARCHAR(50) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    INSERT INTO machine_location (machine, pc_ip, weigher_ip, metal_detector_ip)
    VALUES (@Machine, @PcIp, @WeigherIp, @MetalIp);
END
GO

-- 4. Update existing machine by name
CREATE PROCEDURE [dbo].[sp_UpdateMachineByName]
    @Machine        NVARCHAR(50),
    @PcIp           NVARCHAR(15),
    @WeigherIp      NVARCHAR(15),
    @MetalIp        NVARCHAR(50) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    UPDATE machine_location
    SET pc_ip              = @PcIp,
        weigher_ip         = @WeigherIp,
        metal_detector_ip  = @MetalIp
    WHERE machine = @Machine;
END
GO

-- 5. Delete machine/line
CREATE PROCEDURE [dbo].[sp_DeleteMachine]
    @Machine NVARCHAR(50)
AS
BEGIN
    SET NOCOUNT ON;
    DELETE FROM machine_location WHERE machine = @Machine;
END
GO

-- 6. Fix: only resume an open session — ended sessions always produce a new row
ALTER PROCEDURE [dbo].[sp_GetProductionIdByLot]
    @LotNo   NVARCHAR(50),
    @Machine NVARCHAR(50)
AS
BEGIN
    SET NOCOUNT ON;
    SELECT production_id FROM production_session
    WHERE lot_no = @LotNo AND machine = @Machine AND end_date IS NULL;
END
GO

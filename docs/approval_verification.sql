/* BusyCloud Voucher Approval verification helpers */

-- 1. Approval configuration reserved records
SELECT *
FROM dbo.Config
WHERE RecType IN (203,204)
ORDER BY RecType,[Type],C1;

-- 2. Current approval state + BUSY checklist lifecycle for one voucher
DECLARE @VchCode INT = 0; -- set test code

SELECT
    VchCode,
    VchType,
    VchNo,
    VchSeriesCode,
    ApprovalStatus,
    Status,
    AuditStatus
FROM dbo.Tran1
WHERE VchCode=@VchCode;

SELECT
    Type,
    Code,
    Action,
    ActionTime,
    UserName,
    D1,D2,D3,D4,D5,
    Notes,
    ComputerName
FROM dbo.CheckList
WHERE Code=@VchCode
ORDER BY ActionTime,Action;

-- 3. BusyCloud custom history (table is auto-created by API)
IF OBJECT_ID(N'dbo.BusyCloudVoucherApprovalAudit',N'U') IS NOT NULL
BEGIN
    SELECT TOP 100 *
    FROM dbo.BusyCloudVoucherApprovalAudit
    WHERE VchCode=@VchCode OR @VchCode=0
    ORDER BY ActionTime DESC,Id DESC;
END;

-- 4. Global invariant check
SELECT ApprovalStatus,COUNT(*) AS VoucherCount
FROM dbo.Tran1
GROUP BY ApprovalStatus
ORDER BY ApprovalStatus;

SELECT
    T.ApprovalStatus,
    SUM(CASE WHEN A.Code IS NULL THEN 0 ELSE 1 END) AS WithApprovalAction3,
    COUNT(*) AS VoucherCount
FROM dbo.Tran1 T
LEFT JOIN (
    SELECT DISTINCT Code
    FROM dbo.CheckList
    WHERE Action=3
) A ON A.Code=T.VchCode
GROUP BY T.ApprovalStatus
ORDER BY T.ApprovalStatus;

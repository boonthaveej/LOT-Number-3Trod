/*
USE [LotDB]
GO
/****** Object:  StoredProcedure [dbo].[InsertDataLoop]    Script Date: 08/26/22 10:30:51 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER PROCEDURE [dbo].[InsertDataLoop]
@Lot_ID as bigint,
@Lot_Number as char(1),
@Lot_Val_1 as integer,
@Lot_Val_2 as integer,
@Lot_Date as bigint,
@Login_ID as bigint
AS
BEGIN
DECLARE @i integer
DECLARE @Lot_Flag char(1)
DECLARE @Lot_Status char(2)
SET @i = 0
SET @Lot_Flag = '1'
SET @Lot_Status = '1'
WHILE @i < 10
BEGIN
  INSERT INTO [LotDB].[dbo].[TblLOTDetail] (LOTID, LOTNUMBER, LOTVALUES1, LOTVALUES2, LOTFLAG, LOTTIME, LOTSTATUS, LOTDATE,LoginID)
  VALUES (@Lot_ID,trim(@Lot_Number+convert(char,@i)) ,@Lot_Val_1,@Lot_Val_2,@Lot_Flag,GETDATE(),@Lot_Status,@Lot_Date,@Login_ID)
SET @i = @i + 1
PRINT '>>> [' + trim(@Lot_Number + convert(char,@i)) + '] '
END
-- ###########
SET @i = 9
WHILE @i >-1
BEGIN
if ( trim(convert(char,@i)) <> trim(@Lot_Number))
BEGIN
  INSERT INTO [LotDB].[dbo].[TblLOTDetail] (LOTID, LOTNUMBER, LOTVALUES1, LOTVALUES2, LOTFLAG, LOTTIME, LOTSTATUS, LOTDATE,LoginID)
  VALUES (@Lot_ID,trim(convert(char,@i))+trim(@Lot_Number) ,@Lot_Val_1,@Lot_Val_2,@Lot_Flag,GETDATE(),@Lot_Status,@Lot_Date,@Login_ID)
END
SET  @i = @i -1
PRINT '<<< [' + trim(convert(char,@i)) + trim(@Lot_Number) + '] '
END
END 
*/

/*
USE [LotDB]
GO

/****** Object:  StoredProcedure [dbo].[UpdateConfig]    Script Date: 08/26/22 8:26:43 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

ALTER PROCEDURE [dbo].[UpdateConfig]
        @ConfigTop2Digit decimal,
        @ConfigDown2Digit decimal,
        @ConfigDiv2Digit decimal,
        @ConfigTop3Digit decimal,
        @ConfigDown3Digit decimal,
        @ConfigDiv3Digit decimal,
        @LoginId bigint
        AS
        BEGIN 
		Update [LotDB].[dbo].[TblConfig] SET ConfigTop2Digit  = @ConfigTop2Digit,
        ConfigDown2Digit = @ConfigDown2Digit,
        ConfigDiv2Digit  = @ConfigDiv2Digit,
        ConfigTop3Digit  = @ConfigTop3Digit,
        ConfigDown3Digit = @ConfigDown3Digit,
        ConfigDiv3Digit  = @ConfigDiv3Digit,
        ConfigDate       = GETDATE(),
        LoginID = @LoginId
        WHERE ConfigID = 1
        End
GO


*/
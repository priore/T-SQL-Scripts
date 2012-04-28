--
--  Created by Danilo Priore.
--  Copyright (c) 2012 Prioregroup.com. All rights reserved.
--
--  Permission is hereby granted, free of charge, to any person obtaining a copy
--  of this software and associated documentation files (the "Software"), to deal
--  in the Software without restriction, including without limitation the rights
--  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
--  copies of the Software, and to permit persons to whom the Software is
--  furnished to do so, subject to the following conditions:
--
--  The above copyright notice and this permission notice shall be included in
--  all copies or substantial portions of the Software.
--
--  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
--  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
--  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
--  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
--  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
--  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
--  THE SOFTWARE.
--
CREATE PROCEDURE spConvertAndGetPDF(@filename sysname) AS
BEGIN
SET NOCOUNT ON

-- unique file name for PDF file
DECLARE @pdf varchar(32)
SELECT @pdf = REPLACE(NEWID(),'-','') + '.PDF'

-- converter free tools (Priore StudioPDF+OpenOffice)
-- http://www.prioregroup.com/windows/studiopdf.aspx
EXEC("master..xp_cmdshell 'spdf -in " + @filename + " -out" + @pdf + "', NO_OUTPUT")
IF @@ERROR <> 0 BEGIN
	RETURN
END

-- file size
DECLARE @filesize int
CREATE TABLE #fileinfo(
	fname varchar(255) NULL,
	[size] int NULL,
	unused1 int NULL,
	unused2 int NULL,
	unused3 int NULL,
	unused4 int NULL,
	unused5 int NULL,
	unused6 int NULL,
	unused7 int NULL
)
INSERT #fileinfo exec master..xp_getfiledetails @pdf
SELECT @filesize = [size] FROM #fileinfo
DROP TABLE #fileinfo

-- file not found
IF ISNULL(@pdf,0) = 0 BEGIN
	RETURN
END

-- unique file name for FMT file
DECLARE @dt varchar(23)
SET @dt = CONVERT(varchar(23),GETDATE(),121)
SET @dt = REPLACE(@dt,' ','')
SET @dt = REPLACE(@dt,'-','')
SET @dt = REPLACE(@dt,'/','')
SET @dt = REPLACE(@dt,':','')
SET @dt = REPLACE(@dt,'.','')
SET @dt = REPLACE(@dt,',','')
SET @dt = @dt + '.fmt'

SET @dt = "test.fmt"

-- create FMT file
DECLARE @cmd varchar(200)
SET @cmd = "master..xp_cmdshell '@echo 1 SQLIMAGE 0 " + CONVERT(varchar(10),@filesize)
SET @cmd = @cmd + ' "" 1 img "" >> ' + @dt + "', NO_OUTPUT"
EXEC ("master..xp_cmdshell '@echo 8.0 > " + @dt + "', NO_OUTPUT")
EXEC ("master..xp_cmdshell '@echo 1 >> " + @dt + "', NO_OUTPUT")
EXEC  (@cmd)

-- get and return file
CREATE TABLE #tmpimg (img image)
EXEC("bulk INSERT #tmpimg FROM '" + @pdf + "' WITH (CODEPAGE='RAW',DATAFILETYPE='widenative',FORMATFILE='" + @dt + "')")
IF @@ERROR = 0 BEGIN
	SELECT img FROM #tmpimg
END
ELSE BEGIN
	RETURN
END
DROP TABLE #tmpimg

-- remove temp files
EXEC("master..xp_cmdshell 'del " + @dt + "', NO_OUTPUT")
EXEC("master..xp_cmdshell 'del " + @pdf + "', NO_OUTPUT")

END
GO
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
--  Use:
--  exec spStartDTS "[MY-SERVER\MY-DBSERVER]", "sa", "sa", "my-DTS"
--
CREATE PROCEDURE spStartDTS(
	@server varchar(200), 
	@user varchar(30), 
	@password varchar(30), 
	@name varchar(200)
)
AS
DECLARE @oPKG int
DECLARE @hr int
DECLARE @cmd varchar(100)
SET @cmd = 'LoadFromSQLServer(@server, @user, @password, 256, , , , "[' + @name + ']")'
EXEC @hr = sp_OACreate 'DTS.Package', @oPKG OUT
IF @hr <> 0 BEGIN RETURN @hr END
EXEC @hr = sp_OAMethod @oPKG,@cmd,NULL
IF @hr <> 0 BEGIN RETURN @hr END
EXEC @hr = sp_OAMethod @oPKG, 'Execute'
IF @hr <> 0 BEGIN RETURN @hr END
EXEC @hr = sp_OADestroy @oPKG
RETURN 0
GO
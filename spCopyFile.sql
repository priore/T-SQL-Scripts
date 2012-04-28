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

CREATE PROCEDURE spCopyFile(@source varchar(255), @dest varchar(255))
AS

DECLARE @hr int
DECLARE @ole_FileSystem int
DECLARE @True int
DECLARE @src varchar(250), @desc varchar(2000)

EXEC @hr = sp_OACreate 'Scripting.FileSystemObject', @ole_FileSystem OUT
if @hr <> 0
begin
   exec sp_OAGetErrorInfo @ole_FileSystem, @src OUT, @desc OUT
   raiserror('Object Creation Failed 0x%x, %s, %s',16,1,@hr,@src,@desc)
   return
end

EXEC @hr = sp_OAMethod @ole_FileSystem, 'CopyFile',null, @source, @dest
if @hr <> 0
begin
   exec sp_OAGetErrorInfo @ole_FileSystem, @src OUT, @desc OUT
   exec sp_OADestroy @ole_FileSystem
   raiserror('Method Failed 0x%x, %s, %s',16,1,@hr,@src,@desc)
   return
end


cleanup:
exec @hr = sp_OADestroy @ole_FileSystem
return
GO
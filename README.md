![CLEVER DATA GIT REPO](https://raw.githubusercontent.com/LiCongMingDeShujuku/git-resources/master/0-clever-data-github.png "李聪明的数据库")

# 使用SQL搜索并提取Sharepoint文件
#### Find And Extract Sharepoint Documents With SQL
**发布-日期: 2018年05月11日 (评论)**

![#](images/find-and-extract-sharepoint-dcouments-with-sql-a.png?raw=true "#")

## Contents

- [中文](#中文)
- [English](#English)
- [SQL Logic](#Logic)
- [Build Info](#Build-Info)
- [Author](#Author)
- [License](#License) 


## 中文
这是另一个关于如何使用SQL搜索和提取单个Sharepoint文档的示例。通过上面的图片，你可以看到左侧窗格中我是如何进行基本查询以查找我要查找的文档的。我需要的只是数据库名称，列表名称，文件名和URL，所以我可以绝对定位我要找的文件。我想获得文档创建的时间以及文件的大小这样的额外信息，因此我添加了[alldocs].[timecreated]和[docstreams].[size]。

然后我简单地复制并粘贴了我需要的4个值来提取文件（使用右边的脚本）。

我们可以通过Sharepoint并以这种方式获取文件。文件散布在各种不同的地方。 我创建了两个（查找和提取）脚本来进行该过程，最终我将它们组合成一个大规模自动化程序，可以识别数千个文件，并通过一次单击提取。

我在这个过程中遇到的一个常见问题是，如果你多次运行并忘记删除文件，会发生什么？

希望可以帮助你。 下面是两个脚本。

搜索文件

## English
Here’s another example on how you can both search for, and extract individual Sharepoint Documents with SQL. With the image above you can see how on the left pane I’m doing a basic query to find the documents I’m looking for.
All I needed was the Database Name, List Name, File Name, and the URL so I could positively locate the file. I wanted to get extra information about the file so I added the [alldocs].[timecreated] and [docstreams].[size] just to get an idea of the time when the documents were created, and how large the files are.

I then simply copied and pasted the 4 values I needed to extract the file (using the script on the right).

One could always go through Sharepoint and get the files that way; however these are peppered across a variety of different sites. I created the two (Find & Extract) scripts to sure up the process, but ultimately I combined them together into one massive automation so that many thousands of files could be identified, and extracted by just one click.


One of the more common questions I get around this process is what happens if you run it more than once, and forget to remove the file?

Anyway; hope you find this helpful. Both scripts are below.

SEARCH FOR FILES

---
## Logic
```SQL
use [WSS_Content_Database];
set nocount on
 
select
    'database'  = db_name()
,   'time_created'  = left(alldocs.timecreated, 19)
,   'kb'        = (convert(bigint,alldocstreams.size))/1024
,   'mb'        = (convert(bigint,alldocstreams.size))/1024/1024
,   'list_name' = alllists.tp_title
,   'file_name' = alldocs.leafname
,   'url'       = alldocs.dirname
,   'last_url_folder' = right(alldocs.dirname, charindex('/', reverse('/' + alldocs.dirname)) - 1)
from
    alldocs join alldocstreams  on alldocs.id=alldocstreams.id 
    join alllists           on alllists.tp_id = alldocs.listid
where
    --alldocstreams.[size] > 2048
    right([alldocs].[leafname], 2) in ('oc', 'cx', 'df', 'sg', 'xt')
    and alllists.tp_title like '%FY12 Documents%'
order by
    alldocs.timecreated desc
,   alldocs.dirname 


```

提取文件
EXTRACT THE FILE

```SQL
use master;
set nocount on
 
declare @ole_automation int
set     @ole_automation = (select cast([value_in_use] as int) from sys.configurations where [configuration_id] = '16388')
if      @ole_automation = 0
    begin
    exec sp_configure 'Ole Automation Procedures', 1; reconfigure with override;
    end;
go
 
use tempdb;
set nocount on
 
declare @url            varchar(1000)
declare @list           varchar(255)
declare @file           varchar(255)
declare @database       varchar(255)
declare @extension      varchar(5)
declare @destination_path   varchar(255)
/********************************************************************/
set @database   = 'WSS_Content_Database'
set @list   = 'Archive FY12 Documents'
set @file   = '7684_HiringPacket.pdf'
set @url    = 'sites/Archive of Hiring Docs FY2012'
/********************************************************************/
set @extension = (select reverse(left(reverse(@file),charindex('.',reverse(@file))-1)))
set @destination_path   = '\\sps1\w$\' + @file
 
declare @extract_file   varchar(max)
set @extract_file   = 
'use [' + @database + '];
set nocount on;
 
declare @object_token int
declare @content_binary varbinary(max)
select  @content_binary = alldocstreams.content from alldocs join alldocstreams on alldocs.id = alldocstreams.id join alllists on alllists.tp_id = alldocs.listid
where  
    alllists.tp_title   = ''' + @list + '''
    and alldocs.leafname    = ''' + @file + '''
    and alldocs.dirname = ''' + @url  + '''
 
exec sp_oacreate ''adodb.stream'', @object_token output
exec sp_oasetproperty @object_token, ''type'', 1
exec sp_oamethod @object_token, ''open''
exec sp_oamethod @object_token, ''write'', null, @content_binary
exec sp_oamethod @object_token, ''savetofile'', null, ''' + @destination_path + ''', 2
exec sp_oamethod @object_token, ''close''
exec sp_oadestroy @object_token
'
exec    (@extract_file)

```


[![WorksEveryTime](https://forthebadge.com/images/badges/60-percent-of-the-time-works-every-time.svg)](https://shitday.de/)

## Build-Info

| Build Quality | Build History |
|--|--|
|<table><tr><td>[![Build-Status](https://ci.appveyor.com/api/projects/status/pjxh5g91jpbh7t84?svg?style=flat-square)](#)</td></tr><tr><td>[![Coverage](https://coveralls.io/repos/github/tygerbytes/ResourceFitness/badge.svg?style=flat-square)](#)</td></tr><tr><td>[![Nuget](https://img.shields.io/nuget/v/TW.Resfit.Core.svg?style=flat-square)](#)</td></tr></table>|<table><tr><td>[![Build history](https://buildstats.info/appveyor/chart/tygerbytes/resourcefitness)](#)</td></tr></table>|

## Author

- **李聪明的数据库 Lee's Clever Data**
- **Mike的数据库宝典 Mikes Database Collection**
- **李聪明的数据库** "Lee Songming"

[![Gist](https://img.shields.io/badge/Gist-李聪明的数据库-<COLOR>.svg)](https://gist.github.com/congmingshuju)
[![Twitter](https://img.shields.io/badge/Twitter-mike的数据库宝典-<COLOR>.svg)](https://twitter.com/mikesdatawork?lang=en)
[![Wordpress](https://img.shields.io/badge/Wordpress-mike的数据库宝典-<COLOR>.svg)](https://mikesdatawork.wordpress.com/)

---
## License
[![LicenseCCSA](https://img.shields.io/badge/License-CreativeCommonsSA-<COLOR>.svg)](https://creativecommons.org/share-your-work/licensing-types-examples/)

![Lee Songming](https://raw.githubusercontent.com/LiCongMingDeShujuku/git-resources/master/1-clever-data-github.png "李聪明的数据库")


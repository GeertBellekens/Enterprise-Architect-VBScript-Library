'[path=\Projects\Project A\Temp]
'[group=Temp]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: 
' Author: 
' Purpose: 
' Date: 
'
sub main
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), 'date=2015-1-','date=2015-01-') from t_objectproperties tv where tv.Notes like '%date=2015-1-%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), 'date=2015-2-','date=2015-02-') from t_objectproperties tv where tv.Notes like '%date=2015-2-%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), 'date=2015-3-','date=2015-03-') from t_objectproperties tv where tv.Notes like '%date=2015-3-%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), 'date=2015-4-','date=2015-04-') from t_objectproperties tv where tv.Notes like '%date=2015-4-%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), 'date=2015-5-','date=2015-05-') from t_objectproperties tv where tv.Notes like '%date=2015-5-%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), 'date=2015-6-','date=2015-06-') from t_objectproperties tv where tv.Notes like '%date=2015-6-%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), 'date=2015-7-','date=2015-07-') from t_objectproperties tv where tv.Notes like '%date=2015-7-%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), 'date=2015-8-','date=2015-08-') from t_objectproperties tv where tv.Notes like '%date=2015-8-%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), 'date=2015-9-','date=2015-09-') from t_objectproperties tv where tv.Notes like '%date=2015-9-%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), 'date=2016-1-','date=2016-01-') from t_objectproperties tv where tv.Notes like '%date=2016-1-%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), 'date=2016-2-','date=2016-02-') from t_objectproperties tv where tv.Notes like '%date=2016-2-%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), 'date=2016-3-','date=2016-03-') from t_objectproperties tv where tv.Notes like '%date=2016-3-%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), 'date=2016-4-','date=2016-04-') from t_objectproperties tv where tv.Notes like '%date=2016-4-%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), 'date=2016-5-','date=2016-05-') from t_objectproperties tv where tv.Notes like '%date=2016-5-%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), 'date=2016-6-','date=2016-06-') from t_objectproperties tv where tv.Notes like '%date=2016-6-%'"

	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), 'date=2015-1-','date=2015-01-') from t_attributetag tv where tv.Notes like '%date=2015-1-%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), 'date=2015-2-','date=2015-02-') from t_attributetag tv where tv.Notes like '%date=2015-2-%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), 'date=2015-3-','date=2015-03-') from t_attributetag tv where tv.Notes like '%date=2015-3-%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), 'date=2015-4-','date=2015-04-') from t_attributetag tv where tv.Notes like '%date=2015-4-%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), 'date=2015-5-','date=2015-05-') from t_attributetag tv where tv.Notes like '%date=2015-5-%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), 'date=2015-6-','date=2015-06-') from t_attributetag tv where tv.Notes like '%date=2015-6-%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), 'date=2015-7-','date=2015-07-') from t_attributetag tv where tv.Notes like '%date=2015-7-%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), 'date=2015-8-','date=2015-08-') from t_attributetag tv where tv.Notes like '%date=2015-8-%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), 'date=2015-9-','date=2015-09-') from t_attributetag tv where tv.Notes like '%date=2015-9-%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), 'date=2016-1-','date=2016-01-') from t_attributetag tv where tv.Notes like '%date=2016-1-%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), 'date=2016-2-','date=2016-02-') from t_attributetag tv where tv.Notes like '%date=2016-2-%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), 'date=2016-3-','date=2016-03-') from t_attributetag tv where tv.Notes like '%date=2016-3-%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), 'date=2016-4-','date=2016-04-') from t_attributetag tv where tv.Notes like '%date=2016-4-%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), 'date=2016-5-','date=2016-05-') from t_attributetag tv where tv.Notes like '%date=2016-5-%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), 'date=2016-6-','date=2016-06-') from t_attributetag tv where tv.Notes like '%date=2016-6-%'"

	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), '-1;comments','-01;comments') from t_objectproperties tv where tv.Notes like '%-1;comments%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), '-2;comments','-02;comments') from t_objectproperties tv where tv.Notes like '%-2;comments%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), '-3;comments','-03;comments') from t_objectproperties tv where tv.Notes like '%-3;comments%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), '-4;comments','-04;comments') from t_objectproperties tv where tv.Notes like '%-4;comments%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), '-5;comments','-05;comments') from t_objectproperties tv where tv.Notes like '%-5;comments%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), '-6;comments','-06;comments') from t_objectproperties tv where tv.Notes like '%-6;comments%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), '-7;comments','-07;comments') from t_objectproperties tv where tv.Notes like '%-7;comments%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), '-8;comments','-08;comments') from t_objectproperties tv where tv.Notes like '%-8;comments%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), '-9;comments','-09;comments') from t_objectproperties tv where tv.Notes like '%-9;comments%'"

	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), '-1;comments','-01;comments') from t_attributetag tv where tv.Notes like '%-1;comments%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), '-2;comments','-02;comments') from t_attributetag tv where tv.Notes like '%-2;comments%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), '-3;comments','-03;comments') from t_attributetag tv where tv.Notes like '%-3;comments%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), '-4;comments','-04;comments') from t_attributetag tv where tv.Notes like '%-4;comments%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), '-5;comments','-05;comments') from t_attributetag tv where tv.Notes like '%-5;comments%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), '-6;comments','-06;comments') from t_attributetag tv where tv.Notes like '%-6;comments%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), '-7;comments','-07;comments') from t_attributetag tv where tv.Notes like '%-7;comments%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), '-8;comments','-08;comments') from t_attributetag tv where tv.Notes like '%-8;comments%'"
	Repository.Execute "update tv set tv.Notes = REPLACE(convert(nvarchar(max),tv.Notes), '-9;comments','-09;comments') from t_attributetag tv where tv.Notes like '%-9;comments%'"
end sub

main
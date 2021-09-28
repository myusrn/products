# products
list and search, using regex on name and version, for installed products from a command line interface

status = in progress  
version = 28sep21

28sep21 recompiling using vs22 and nfx48 and fixed exception arising iterating through Product.Mu where Patches subkey does not exist
05feb18 recompiling made Products.Mu /t all processing exception surfacing in new os environments go away  
27dec17 added /t nonmsi processing that looks in CurrentUser registry path in addition to LocalMachine registry path  
02jul12 added /x processing use of ignoredependencies=all property that causes silent uninstalls when msi has chained ref counts  
05mar12 switch /t upd[ate] to /t mu and /t wu to allow differentiating between microsoft update and windows update bits  
29feb12 added /t upd[ate] support and not sure what change change you made on 02dec10 but it appears ProductsMsiRcw.dll is no longer needed  
30sep11 added /ss, /r, /rs, /rv switch support to facilitate addRmvPrd.cmd script processing w/o needing pids for everything  
29jul11 add /vsl switch to allow /v output to have \n carriage return and /vsl to not to faciliate cases when you need to psh short and diff results  
29jun11 removed \n carriage return in verbose output cases to make psh sorting of results dumped into text file work  
02dec10 tried to change ProductsMsiRcw.dll dependency to an embedded resource but appears more work to do that so just updated it to be current  
24feb09 migrated to vs10 and reviewed ability to remove runtime callable [com] wrapper dependency by switching to use of managed wmi calls  
04aug06 attempted to enable unc access by using strong name key signing and tlbimp /unsafe switch  
23apr06 added x64 and nonMsiX86 support  
14mar06 migrated to vs05 and added optional x64 build support  
01mar04 created 1st drop  
  
Products Command Line Interface Utility
------------------------------------------------------------------------------------------------------------

this utitlity is intended to enable a command line solution for reviewing installed msi or nonmsi products and updates installed on a given system

for listing windows update [ / hotfixes ] there also exists
- out of the box systeminfo cli which lists them but with no additional output details such as date and time of install and total hits or search capabilities
- out of the box dism /online /get-packages which has similar output details missing and lists items with visible = 1 and visible = 0 
- sysinternals psinfo -h which has similar output details missing appears to currently require some updating in order to produce matching results

TODO: consider update to include /sp that allows you to search by pid value 

TODO: optinally update to remove embedded resource interop dll dependency [ as outlined in silverlight openFileDialog com interop sample ]
and instead use wmi root\cimv2 select * from win32_product

TODO: look at windows sdk samples\sysmgmt\msi\scripts\WiLstPrd.vbs and leverage routines it uses for listing installed products 
"cscript WiLstPrd.vbs" and listing product installed components "cscript WiLstPrd.vbs <installed product guid> pcd"
  
test commands
------------------------------------------------------------------------------------------------------------
/s "Silverlight \\d \\w* ?SDK"
/ss "src\\ofc10" /vsl
/ss "src\\vs12\\x86" /vsl
/s "v11.0.5\d{4}\\packages" /vsl
/ss "v11.0.5\d{4}\\packages" /vsl
/ss "src\\vs12\\x86|11.0.5\d{4}"
/sv "11.0.4\d{4}" /vsl
/s "Microsoft .NET Framework (?:4.5|4.5 Developer Preview) Multi-Targeting Pack"
/r "Blend 5" /vsl
/l /t msi /vsl
/l /t nonmsi /vsl
/l /t mu
/l /t mu /vsl
/s lync /t mu
/s lync /t mu /vsl
/l /t wu
/l /t wu /vsl
/s lync /t wu
/s lync /t wu /vsl
  
interop file creation 
------------------------------------------------------------------------------------------------------------
pushd d:\sswf\ops\products  
attrib -r *Rcw.dll & del /f *Rcw.dll  
set tlbImpExe=%WindowsSDK_ExecutablePath_x86%\tlbimp.exe  
set comLibPath=%systemroot%\system32& set machine=x64& set namespacePrefix=MyCompany.Ops.Utils.Products  
//"%tlbImpExe%" "%comLibPath%\msi.dll" /out:"ProductsMsiRcw.dll" /namespace:%namespacePrefix%.MsiRcw /keyfile:"Products.snk" /machine:%machine% /unsafe  
"%tlbImpExe%" "%comLibPath%\msi.dll" /out:"ProductsMsiRcw.dll" /namespace:%namespacePrefix%.MsiRcw /keyfile:"Products.snk"  
add reference to rcw and set "Embed Interop Types" = true  
set namespacePrefix=&set machine=&set comLibPath=&set tlbImpExe=  
note1 - don't include machine switch if you want to create x86 and x64 compatible interop library output for types where this is possible, i.e. when the com library's x86 and x64 clsids and progIds are the same  
note2 - marking interop library as /unsafe did not allow partially trusted callers, e.g. operation from a unc path, as it was suggested it would but now it just works from unc paths so perhaps something in nfx20/35/4x processing has changed  
  
debug tested command line arguments  
------------------------------------------------------------------------------------------------------------
&lt;noargs&gt;  
/? and /h  
/t msi /l  
/t nonmsi /l /v  
/t msi /s office  
/t nonmsi /s sql /v  
/t msi /s " 2003"  
 

Attribute VB_Name = "VBAACI"
Option Explicit

'Returns a token to be used with the rest of the functions in this module
Function Authenticate(url, Username, Password) As String
Dim token As String, aux As String, i
        'Create Web Client
        Dim ACIclient As New WebClient
        'If the variable wasnt initialized, open the form so that the user enters credentials
        If Not Len(url) > 0 Then CredentialsForm.Show
        'If the variable is still null, the use clicked on cancel
        If Not Len(url) > 0 Then Exit Function
        ACIclient.BaseUrl = url
        'Accept self-signed certificates
        ACIclient.Insecure = True
        'Create Request
        Dim ACIrequest As New WebRequest
        ACIrequest.Resource = "aaaLogin.json"
        ACIrequest.method = WebMethod.HttpPost
        ACIrequest.Body = "{""aaaUser"" : {""attributes"" : {""name"" : """ & Username & """, ""pwd"" : """ & Password & """}}}"
        ACIrequest.SetHeader "Content-Type", "application/json"
        'Send request and store response
        Dim Response As WebResponse
        Set Response = ACIclient.Execute(ACIrequest)
        If Not Response.StatusCode = 200 Then
            MsgBox "Error in authentication: " & Response.StatusCode & " - " & Response.StatusDescription
            Exit Function
        End If
        'Get cookie
        Dim Json As Object
        Set Json = JsonConverter.ParseJson(Response.Content)
        token = JsonConverter.ConvertToJson(Json("imdata").Item(1)("aaaLogin")("attributes")("token"))
        'Strip double quotes from response (first and last character)
        token = Mid(token, 2, Len(token) - 2)
        'Remove inverted slash from response
        aux = ""
        For i = 1 To Len(token)
            If Mid(token, i, 1) <> "\" Then aux = aux & Mid(token, i, 1)
        Next i
        Authenticate = aux
End Function

'Send a generic POST, return True if successful
Function SendACIPOST(token As String, Resource As String, Body As String, errmsg As String) As Boolean
Dim Name As String, Id As String, JSONpreview As Boolean
        'Construct Web Client
        Dim ACIclient As New WebClient
        ACIclient.BaseUrl = apic_url
        ACIclient.Insecure = True
        'Construct request
        Dim ACIrequest As New WebRequest
        ACIrequest.Resource = Resource
        ACIrequest.method = WebMethod.HttpPost
        'Construct Body
        ACIrequest.Body = Body
        ACIrequest.SetHeader "Cookie", "APIC-cookie=" & token
        ACIrequest.SetHeader "Content-Type", "application/json"
        'Send request
        Dim Response As WebResponse
        Set Response = ACIclient.Execute(ACIrequest)
        If Not Response.StatusCode = 200 Then
            SendACIPOST = False
        Else
            SendACIPOST = True
        End If
End Function

'Generic GET call to the REST API that checks the totalCount field of the JSON return
Function ObjectExists(token As String, Resource As String) As Boolean
        'Construct Web Client
        Dim ACIclient As New WebClient
        ACIclient.BaseUrl = apic_url
        ACIclient.Insecure = True
        'Construct request
        Dim ACIrequest As New WebRequest
        ACIrequest.Resource = Resource
        ACIrequest.method = WebMethod.HttpGet
        ACIrequest.SetHeader "Cookie", "APIC-cookie=" & token
        ACIrequest.SetHeader "Content-Type", "application/json"
        'Send request
        Dim Response As New WebResponse
        Set Response = ACIclient.Execute(ACIrequest)
        If Not Response.StatusCode = 200 Then
            MsgBox "Error checking for object: " & Response.StatusCode & " - " & Response.StatusDescription, vbCritical
        Else
            'Process answer
            Dim Json As Object
            Set Json = JsonConverter.ParseJson(Response.Content)
            ObjectExists = (Mid(JsonConverter.ConvertToJson(Json("totalCount")), 2, 1) > 0)
        End If
End Function

Function DeleteIfSelector(token As String, intprofile As String, servername As String)
'url: http://muc-apic/api/node/mo/uni/infra/accportprof-201-202.json
'payload: {"infraAccPortP":{"attributes":{"dn":"uni/infra/accportprof-201-202","status":"modified"},"children":[{"infraHPortS":{"attributes":{"dn":"uni/infra/accportprof-201-202/hports-server-0134-typ-range","status":"deleted"},"children":[]}}]}}
 Dim Resource As String, errmsg As String, Body As String
        Resource = "node/mo/uni/infra/accportprof-" & intprofile & ".json"
        Body = "{""infraAccPortP"":{""attributes"":{""dn"":""uni/infra/accportprof-" & intprofile & """,""status"":""modified""}," _
        & """children"":[{""infraHPortS"":{""attributes"":{""dn"":""uni/infra/accportprof-" & intprofile & "/hports-" & servername & "-typ-range""," _
        & """status"":""deleted""},""children"":[]}}]}}"
        errmsg = "Error in VPC policy group delete request"
        SendPOST token, Resource, Body, errmsg
End Function

Function DeleteIfSelectorFEX(token As String, intprofile As String, servername As String)
'url: http://muc-apic/api/node/mo/uni/infra/fexprof-101.json
'payload: {"infraFexP":{"attributes":{"dn":"uni/infra/fexprof-101","status":"modified"},"children":[{"infraHPortS":{"attributes":{"dn":"uni/infra/fexprof-101/hports-test_fex-typ-range","status":"deleted"},"children":[]}}]}}
'payload: {"infraFexP":{"attributes":{"dn":"uni/infra/fexprof-102","status":"modified"},"children":[{"infraHPortS":{"attributes":{"dn":"uni/infra/fexprof-102/hports-server-1133-typ-range","status":"deleted"},"children":[]}}]}}
'debug:   {"infraFexP":{"attributes":{"dn":"uni/infra/fexprof-101","status":"modified"},"children":[{"infraHPortS":{"attributes":{"dn":"uni/infra/fexprof-101/hports-server-1133-typ-range","status":"deleted"},"children":[]}}]}}
Dim Resource As String, errmsg As String, Body As String
        Resource = "node/mo/uni/infra/fexprof-" & intprofile & ".json"
        Body = "{""infraFexP"":{""attributes"":{""dn"":""uni/infra/fexprof-" & intprofile & """," _
        & """status"":""modified""},""children"":[{""infraHPortS"":{""attributes"":{""dn"":""uni/infra/fexprof-" & intprofile & "/hports-" & servername & "-typ-range""," _
        & """status"":""deleted""},""children"":[]}}]}}"
        errmsg = "Error deleting FEX ifselector " & servername & " from FEX " & intprofile
        SendPOST token, Resource, Body, errmsg
End Function

Function ConfigurePort(token As String, intprofile As String, servername As String, port As String, policy As String)
Dim Resource As String, errmsg As String, Body As String
        Resource = "node/mo/uni/infra/accportprof-" & intprofile & "/hports-test-typ-range.json"
        Body = "{""infraHPortS"":{""attributes"":{""dn"":""uni/infra/accportprof-" & intprofile & "/hports-" & servername & "-typ-range""," _
        & """name"":""" & servername & """,""rn"":""hports-" & servername & "-typ-range"",""status"":""created,modified""},""children"":[" _
        & "{""infraPortBlk"":{""attributes"":{""dn"":""uni/infra/accportprof-" & intprofile & "/hports-" & servername & "-typ-range/portblk-block2""," _
        & """fromPort"":""" & port & """,""toPort"":""" & port & """,""name"":""block2"",""rn"":""portblk-block2"",""status"":""created,modified""}," _
        & """children"":[]}},{""infraRsAccBaseGrp"":{""attributes"":{""tDn"":""uni/infra/funcprof/accportgrp-" & policy & """,""status"":""created,modified""},""children"":[]}}]}}"
        errmsg = "Error in configuration request"
        SendPOST token, Resource, Body, errmsg
End Function

Function ConfigureFEXPortVPC(token As String, intprofile As String, servername As String, port As String, policy As String)
'Identical to "ConfigureFEXPort", but we need to use "accbundle" instead of "accportgrp"
Dim Resource As String, errmsg As String, Body As String
        Resource = "node/mo/uni/infra/fexprof-" & intprofile & "/hports-" & servername & "-typ-range.json"
        Body = "{""infraHPortS"":{""attributes"":{""dn"":""uni/infra/fexprof-" & intprofile & "/hports-" & servername & "-typ-range""," _
        & """name"":""" & servername & """,""rn"":""hports-" & servername & "-typ-range"",""status"":""created,modified""},""children"":[{""infraPortBlk"":" _
        & "{""attributes"":{""dn"":""uni/infra/fexprof-" & intprofile & "/hports-" & servername & "-typ-range/portblk-block3"",""fromPort"":""" & port & """," _
        & """toPort"":""" & port & """,""name"":""block3"",""rn"":""portblk-block3"",""status"":""created,modified""},""children"":[]}},{""infraRsAccBaseGrp"":" _
        & "{""attributes"":{""tDn"":""uni/infra/funcprof/accbundle-" & policy & """,""status"":""created,modified""},""children"":[]}}]}}"
        errmsg = "Error in request to add ifselector for " & servername & " on FEX " & intprofile & ", port " & port
        SendPOST token, Resource, Body, errmsg
End Function

Function ConfigureFEXPort(token As String, intprofile As String, servername As String, port As String, policy As String)
'url: http://muc-apic/api/node/mo/uni/infra/fexprof-101/hports-test_fex-typ-range.json
'payload: {"infraHPortS":{"attributes":{"dn":"uni/infra/fexprof-101/hports-test_fex-typ-range","name":"test_fex","rn":"hports-test_fex-typ-range","status":"created,modified"},"children":[{"infraPortBlk":{"attributes":{"dn":"uni/infra/fexprof-101/hports-test_fex-typ-range/portblk-block3","fromPort":"30","toPort":"30","name":"block3","rn":"portblk-block3","status":"created,modified"},"children":[]}},{"infraRsAccBaseGrp":{"attributes":{"tDn":"uni/infra/funcprof/accportgrp-1G","status":"created,modified"},"children":[]}}]}}
Dim Resource As String, errmsg As String, Body As String
        Resource = "node/mo/uni/infra/fexprof-" & intprofile & "/hports-" & servername & "-typ-range.json"
        Body = "{""infraHPortS"":{""attributes"":{""dn"":""uni/infra/fexprof-" & intprofile & "/hports-" & servername & "-typ-range""," _
        & """name"":""" & servername & """,""rn"":""hports-" & servername & "-typ-range"",""status"":""created,modified""},""children"":[{""infraPortBlk"":" _
        & "{""attributes"":{""dn"":""uni/infra/fexprof-" & intprofile & "/hports-" & servername & "-typ-range/portblk-block3"",""fromPort"":""" & port & """," _
        & """toPort"":""" & port & """,""name"":""block3"",""rn"":""portblk-block3"",""status"":""created,modified""},""children"":[]}},{""infraRsAccBaseGrp"":" _
        & "{""attributes"":{""tDn"":""uni/infra/funcprof/accportgrp-" & policy & """,""status"":""created,modified""},""children"":[]}}]}}"
        errmsg = "Error in request to add ifselector for " & servername & " on FEX " & intprofile & ", port " & port
        SendPOST token, Resource, Body, errmsg
End Function


Function AddChannelPolicy(token As String, servername As String)
Dim Resource As String, errmsg As String, Body As String
        Resource = "node/mo/uni/infra/lacplagp-" & servername & ".json"
        Body = "{""lacpLagPol"":{""attributes"":{""dn"":""uni/infra/lacplagp-" & servername & """,""ctrl"":""fast" _
        & "-sel-hot-stdby,graceful-conv,susp-individual"",""name"":""" & servername & """,""mode"":""active"",""rn"":""" _
        & "lacplagp-" & servername & """,""status"":""created""},""children"":[]}}"
        errmsg = "Error in PC policy request"
        SendPOST token, Resource, Body, errmsg
End Function

Function DeleteChannelPolicy(token As String, servername As String)
'url: http://muc-apic/api/node/mo/uni/infra.json
'payload: {"infraInfra":{"attributes":{"dn":"uni/infra","status":"modified"},"children":[{"lacpLagPol":{"attributes":{"dn":"uni/infra/lacplagp-server-0134","status":"deleted"},"children":[]}}]}}
Dim Resource As String, errmsg As String, Body As String
        Resource = "node/mo/uni/infra.json"
        Body = "{""infraInfra"":{""attributes"":{""dn"":""uni/infra"",""status"":""modified""},""children"":[{""lacpLagPol"":" _
        & "{""attributes"":{""dn"":""uni/infra/lacplagp-" & servername & """,""status"":""deleted""},""children"":[]}}]}}"
        errmsg = "Error in PC policy delete request"
        SendPOST token, Resource, Body, errmsg
End Function

Function DeleteVPCPolicyGroup(token As String, servername As String)
'url: http://muc-apic/api/node/mo/uni/infra/funcprof.json
'payload: {"infraFuncP":{"attributes":{"dn":"uni/infra/funcprof","status":"modified"},"children":[{"infraAccBndlGrp":{"attributes":{"dn":"uni/infra/funcprof/accbundle-server-0134","status":"deleted"},"children":[]}}]}}
Dim Resource As String, errmsg As String, Body As String
        Resource = "node/mo/uni/infra/funcprof.json"
        Body = "{""infraFuncP"":{""attributes"":{""dn"":""uni/infra/funcprof"",""status"":""modified""},""children"":[" _
        & "{""infraAccBndlGrp"":{""attributes"":{""dn"":""uni/infra/funcprof/accbundle-" & servername & """,""status"":""deleted""},""children"":[]}}]}}"
        errmsg = "Error in VPC policy group delete request"
        SendPOST token, Resource, Body, errmsg
End Function

Function AddVPCPolicyGroup(token As String, servername As String, LLPolicy As String, aep As String)
Dim Resource As String, errmsg As String, Body As String
        Resource = "node/mo/uni/infra/funcprof/accbundle-" & servername & ".json"
        Body = "{""infraAccBndlGrp"":{""attributes"":{""dn"":""uni/infra/funcprof/accbundle-" & servername & """" _
        & ",""lagT"":""node"",""name"":""" & servername & """,""rn"":""accbundle-" & servername & """,""status"":""created""}" _
        & ",""children"":[{""infraRsAttEntP"":{""attributes"":{""tDn"":""uni/infra/attentp-" & aep & """,""status"":""created,modified""}" _
        & ",""children"":[]}},{""infraRsHIfPol"":{""attributes"":{""tnFabricHIfPolName"":""" & LLPolicy & """,""status"":""created,modified""}" _
        & ",""children"":[]}},{""infraRsCdpIfPol"":{""attributes"":{""tnCdpIfPolName"":""CDP-ON"",""status"":""created,modified""}" _
        & ",""children"":[]}},{""infraRsLldpIfPol"":{""attributes"":{""tnLldpIfPolName"":""LLDP-ON"",""status"":""created,modified""}" _
        & ",""children"":[]}},{""infraRsL2IfPol"":{""attributes"":{""tnL2IfPolName"":""default"",""status"":""created,modified""},""children"":[]}}" _
        & ",{""infraRsLacpPol"":{""attributes"":{""tnLacpLagPolName"":""" & servername & """,""status"":""created,modified""},""children"":[]}}]}}"
        errmsg = "Error in VPC policy group request"
        SendPOST token, Resource, Body, errmsg
End Function

Function AddStaticBinding(token As String, tenant As String, anp As String, epg As String, switch, port, vlanid, native As Boolean)
'url: http://muc-apic/api/node/mo/uni/tn-Acme/ap-MyApp1/epg-Tier1.json
'Tagged:
'payload: {"fvRsPathAtt":{"attributes":{"encap":"vlan-1098","instrImedcy":"immediate","tDn":"topology/pod-1/paths-201/pathep-[eth1/15]","status":"created"},"children":[]}}
'Untagged:
'payload: {"fvRsPathAtt":{"attributes":{"encap":"vlan-1097","mode":"untagged","tDn":"topology/pod-1/paths-201/pathep-[eth1/18]","status":"created"},"children":[]}}
Dim Resource As String, errmsg As String, Body As String
        Resource = "node/mo/uni/tn-" & tenant & "/ap-" & anp & "/epg-" & epg & ".json"
        Body = "{""fvRsPathAtt"":{""attributes"":{""encap"":""vlan-" & vlanid & """,""instrImedcy"":""immediate""," _
        & IIf(native, """mode"":""native"", ", "") _
        & """tDn"":""topology/pod-1/paths-" & switch & "/pathep-[eth1/" & port & "]"",""status"":""created""},""children"":[]}}"
        errmsg = "Error in static binding request"
        SendPOST token, Resource, Body, errmsg
End Function

Function AddStaticBindingFEX(token As String, tenant As String, anp As String, epg As String, fex, switch, port, vlanid, native As Boolean)
'url: http://muc-apic/api/node/mo/uni/tn-Acme/ap-Backup/epg-Backup.json
'payload: {"fvRsPathAtt":{"attributes":{"encap":"vlan-1118","instrImedcy":"immediate","tDn":
'               "topology/pod-1/paths-201/extpaths-101/pathep-[eth1/14]","status":"created"},"children":[]}}
Dim Resource As String, errmsg As String, Body As String
        Resource = "node/mo/uni/tn-" & tenant & "/ap-" & anp & "/epg-" & epg & ".json"
        Body = "{""fvRsPathAtt"":{""attributes"":{""encap"":""vlan-" & vlanid & """,""instrImedcy"":""immediate""," _
        & IIf(native, """mode"":""native"", ", "") _
        & """tDn"":""topology/pod-1/paths-" & switch & "/extpaths-" & fex & "/pathep-[eth1/" & port & "]"",""status"":""created""},""children"":[]}}"
        errmsg = "Error in static FEX binding request to " & tenant & "/" & anp & "/" & epg & " on port " & switch & "/" & fex & "/" & port & " with VLAN " & vlanid
        SendPOST token, Resource, Body, errmsg
End Function

Function DeleteStaticBinding(token As String, tenant As String, anp As String, epg As String, switch, port)
'url: http://muc-apic/api/node/mo/uni/tn-Acme/ap-MyApp1/epg-Tier1.json
'payload: {"fvAEPg":{"attributes":{"dn":"uni/tn-Acme/ap-MyApp1/epg-Tier1","status":"modified"},"children":[{"fvRsPathAtt":{"attributes":{"dn":"uni/tn-Acme/ap-MyApp1/epg-Tier1/rspathAtt-[topology/pod-1/paths-201/pathep-[eth1/15\]","status":"deleted"},"children":[]}}]}}
Dim Resource As String, errmsg As String, Body As String
        Resource = "node/mo/uni/tn-" & tenant & "/ap-" & anp & "/epg-" & epg & ".json"
        Body = "{""fvAEPg"":{""attributes"":{""dn"":""uni/tn-" & tenant & "/ap-" & anp & "/epg-" & epg & """,""status"":""modified""}," _
        & """children"":[{""fvRsPathAtt"":{""attributes"":{""dn"":""uni/tn-" & tenant & "/ap-" & anp & "/epg-" & epg & "/rspathAtt-" _
        & "[topology/pod-1/paths-" & switch & "/pathep-[eth1/" & port & "]]"",""status"":""deleted""},""children"":[]}}]}}"
        errmsg = "Error in static binding delete request"
        SendPOST token, Resource, Body, errmsg
End Function


Function DeleteStaticBindingFEX(token As String, tenant As String, anp As String, epg As String, switch, fex, port)
'url: http://muc-apic/api/node/mo/uni/tn-Acme/ap-Backup/epg-Backup.json
'payload: {"fvAEPg":{"attributes":{"dn":"uni/tn-Acme/ap-Backup/epg-Backup","status":"modified"},"children":[{"fvRsPathAtt":{"attributes":{"dn":"uni/tn-Acme/ap-Backup/epg-Backup/rspathAtt-
'                 [topology/pod-1/paths-201/extpaths-101/pathep-[eth1/14\]","status":"deleted"},"children":[]}}]}}
Dim Resource As String, errmsg As String, Body As String
        Resource = "node/mo/uni/tn-" & tenant & "/ap-" & anp & "/epg-" & epg & ".json"
        Body = "{""fvAEPg"":{""attributes"":{""dn"":""uni/tn-" & tenant & "/ap-" & anp & "/epg-" & epg & """,""status"":""modified""}," _
        & """children"":[{""fvRsPathAtt"":{""attributes"":{""dn"":""uni/tn-" & tenant & "/ap-" & anp & "/epg-" & epg & "/rspathAtt-" _
        & "[topology/pod-1/paths-" & switch & "/extpaths-" & fex & "/pathep-[eth1/" & port & "]]"",""status"":""deleted""},""children"":[]}}]}}"
        errmsg = "Error in FEX static binding delete request to " & tenant & "/" & anp & "/" & epg & " on port " & switch & "/" & fex & "/" & port
        SendPOST token, Resource, Body, errmsg
End Function

Function AddStaticBindingVPC(token As String, tenant As String, anp As String, epg As String, intprofile As String, servername As String, vlanid As String, native As Boolean)
'url: http://muc-apic/api/node/mo/uni/tn-Acme/ap-MyApp1/epg-Tier1.json
'payload: {"fvRsPathAtt":{"attributes":{"encap":"vlan-1045","instrImedcy":"immediate",
'"tDn":"topology/pod-1/protpaths-201-202/pathep-[UCS-FI-A]","status":"created"},"children":[]}}
Dim Resource As String, errmsg As String, Body As String
        Resource = "node/mo/uni/tn-" & tenant & "/ap-" & anp & "/epg-" & epg & ".json"
        Body = "{""fvRsPathAtt"":{""attributes"":{""encap"":""vlan-" & vlanid & """,""instrImedcy"":""immediate""," _
        & IIf(native, """mode"":""native""", "") _
        & """tDn"":""topology/pod-1/protpaths-" & intprofile & "/pathep-[" & servername & "]"",""status"":""created""},""children"":[]}}"
        errmsg = "Error in static VPC binding request"
        SendPOST token, Resource, Body, errmsg
End Function

Function AddStaticBindingFEXVPC(token As String, tenant As String, anp As String, epg As String, intprofile As String, parentprofile As String, servername As String, vlanid As String, native As Boolean)
'url: https://muc-apic.cisco.com/api/node/mo/uni/tn-Storage/ap-iSCSI/epg-Targets.json
'payload{"fvRsPathAtt":{"attributes":{"encap":"vlan-1089","instrImedcy":"immediate","mode":"native","tDn":"topology/pod-1/protpaths-201-202/extprotpaths-101-102/pathep-[iSCSI-b]","status":"created"},"children":[]}}
Dim Resource As String, errmsg As String, Body As String
        Resource = "node/mo/uni/tn-" & tenant & "/ap-" & anp & "/epg-" & epg & ".json"
        Body = "{""fvRsPathAtt"":{""attributes"":{""encap"":""vlan-" & vlanid & """,""instrImedcy"":""immediate""," _
        & IIf(native, """mode"":""native""", "") _
        & """tDn"":""topology/pod-1/protpaths-" & parentprofile & "/extprotpaths-" & intprofile & "/pathep-[" & servername & "]"",""status"":""created""},""children"":[]}}"
        errmsg = "Error in static VPC binding request for server " & servername & " on EPG " & epg
        SendPOST token, Resource, Body, errmsg
End Function

Function DeleteStaticBindingVPC(token As String, tenant As String, anp As String, epg As String, intprofile, servername)
'url: http://muc-apic/api/node/mo/uni/tn-Acme/ap-MyApp1/epg-Tier1.json
'payload: {"fvAEPg":{"attributes":{"dn":"uni/tn-Acme/ap-MyApp1/epg-Tier1","status":"modified"},"children":
'[{"fvRsPathAtt":{"attributes":{"dn":"uni/tn-Acme/ap-MyApp1/epg-Tier1/rspathAtt-[topology/pod-1/
'protpaths-201-202/pathep-[UCS-FI-A\]","status":"deleted"},"children":[]}}]}}
Dim Resource As String, errmsg As String, Body As String
        Resource = "node/mo/uni/tn-" & tenant & "/ap-" & anp & "/epg-" & epg & ".json"
        Body = "{""fvAEPg"":{""attributes"":{""dn"":""uni/tn-" & tenant & "/ap-" & anp & "/epg-" & epg & """,""status"":""modified""}," _
        & """children"":[{""fvRsPathAtt"":{""attributes"":{""dn"":""uni/tn-" & tenant & "/ap-" & anp & "/epg-" & epg & "/rspathAtt-" _
        & "[topology/pod-1/protpaths-" & intprofile & "/pathep-[" & servername & "]]"",""status"":""deleted""},""children"":[]}}]}}"
        errmsg = "Error in static binding delete request"
        SendPOST token, Resource, Body, errmsg
End Function

Function CreateSwitchPolicyGroup(token As String, polname As String)
'url: http://muc-apic/api/node/mo/uni/infra/funcprof/accnodepgrp-allleafs.json
'payload: {"infraAccNodePGrp":{"attributes":{"dn":"uni/infra/funcprof/accnodepgrp-allleafs","name":"allleafs","rn":"accnodepgrp-allleafs",
'"status":"created"},"children":[{"infraRsMstInstPol":{"attributes":{"tnStpInstPolName":"default","status":"created,modified"},"children":[]}},
'{"infraRsMonNodeInfraPol":{"attributes":{"tnMonInfraPolName":"default","status":"created,modified"},"children":[]}}]}}
Dim Resource As String, errmsg As String, Body As String
       Resource = "node/mo/uni/infra/funcprof/accnodepgrp-" & polname & ".json"
       Body = "{""infraAccNodePGrp"":{""attributes"":{""dn"":""uni/infra/funcprof/accnodepgrp-" & polname & """,""name"":""" & polname & """," _
        & """rn"":""accnodepgrp-" & polname & """,""status"":""created,modified""},""children"":[{""infraRsMstInstPol"":{""attributes"":{""tnStpInstPolName"":""default""," _
        & """status"":""created,modified""},""children"":[]}},{""infraRsMonNodeInfraPol"":{""attributes"":{""tnMonInfraPolName"":""default"",""status"":""created,modified""}" _
        & ",""children"":[]}}]}}"
       errmsg = "Error creating switch policy"
       SendPOST token, Resource, Body, errmsg
End Function

'This function creates a switch profile including 2 switches, associated to a policy
Function CreateSwitchProfile(token As String, profname As String, switch1, switch2, policy)
'method: Post
'url: http://muc-apic/api/node/mo/uni/infra.json
'payload: {"infraNodeP":{"attributes":{"descr":"This is a test","dn":"uni/infra/nprof-mytest","name":"mytest","ownerKey":"","ownerTag":""},"children":[{"infraLeafS":{"attributes":{"descr":"","name":"201-202","ownerKey":"","ownerTag":"","type":"range"},"children":[{"infraNodeBlk":{"attributes":{"descr":"","from_":"201","name":"a061b0ef8cd1d87b","to_":"202"}}},{"infraRsAccNodePGrp":{"attributes":{"tDn":"uni/infra/funcprof/accnodepgrp-Leafs"}}}]}},{"infraRsAccPortP":{"attributes":{"tDn":"uni/infra/accportprof-201-202"}}}]}}
Dim Resource As String, errmsg As String, Body As String
       Resource = "node/mo/uni/infra.json"
       Body = "{""infraNodeP"":{""attributes"":{""descr"":""Switch Profile"",""dn"":""uni/infra/nprof-" & profname & """,""name"":""" & profname & """," _
        & """ownerKey"":"""",""ownerTag"":""""},""children"":[{""infraLeafS"":{""attributes"":{""descr"":""Switch selector"",""name"":""" & profname & """,""ownerKey"":""""," _
        & """ownerTag"":"""",""type"":""range""},""children"":[{""infraNodeBlk"":{""attributes"":{""descr"":"""",""from_"":""" & switch1 & """,""name"":""a061b0ef8cd1d87b""," _
        & """to_"":""" & switch2 & """}}},{""infraRsAccNodePGrp"":{""attributes"":{""tDn"":""uni/infra/funcprof/accnodepgrp-" & policy & """}}}]}}," _
        & "{""infraRsAccPortP"":{""attributes"":{""tDn"":""uni/infra/accportprof-" & profname & """}}}]}}"
       errmsg = "Error creating switch profile"
       SendPOST token, Resource, Body, errmsg
End Function

Function CreateIntProfile(token As String, profname As String)
'url: http://muc-apic/api/node/mo/uni/infra/accportprof-test.json
'payload: {"infraAccPortP":{"attributes":{"dn":"uni/infra/accportprof-test","name":"test","rn":"accportprof-test","status":"created,modified"},"children":[]}}
Dim Resource As String, errmsg As String, Body As String
       Resource = "node/mo/uni/infra/accportprof-" & profname & ".json"
       Body = "{""infraAccPortP"":{""attributes"":{""dn"":""uni/infra/accportprof-" & profname & """,""name"":""" & profname & """,""rn"":""accportprof-" & profname & """," _
        & """status"":""created,modified""},""children"":[]}}"
       errmsg = "Error creating interfaceprofile"
       SendPOST token, Resource, Body, errmsg
End Function

Function CreateTenant(token As String, tenantname As String)
'url:Êhttp://muc-apic/api/node/mo/uni/tn-testtenant.json
'payload:Ê{"fvTenant":{"attributes":{"dn":"uni/tn-testtenant","name":"testtenant","rn":"tn-testtenant","status":"created"},"children":[]}}
Dim Resource As String, errmsg As String, Body As String
        Resource = "node/mo/uni/tn-" & tenantname & ".json"
        Body = "{""fvTenant"":{""attributes"":{""dn"":""uni/tn-" & tenantname & """,""name"":""" & tenantname & """,""rn"":""tn-" & tenantname & """,""status"":""created""},""children"":[]}}"
        errmsg = "Error creating tenant " & tenantname
        SendPOST token, Resource, Body, errmsg
End Function

Function CreateANP(token As String, tenant As String, anp As String)
'url:Êhttp://muc-apic/api/node/mo/uni/tn-testtenant/ap-testanp.json
'payload:Ê{"fvAp":{"attributes":{"dn":"uni/tn-testtenant/ap-testanp","name":"testanp","rn":"ap-testanp","status":"created"},"children":[]}}
Dim Resource As String, errmsg As String, Body As String
        Resource = "node/mo/uni/tn-" & tenant & "/ap-" & anp & ".json"
        Body = "{""fvAp"":{""attributes"":{""dn"":""uni/tn-" & tenant & "/ap-" & anp & """,""name"":""" & anp & """,""rn"":""ap-" & anp & """,""status"":""created""},""children"":[]}}"
        errmsg = "Error creating ANP " & anp & " in tenant " & tenant
        SendPOST token, Resource, Body, errmsg
End Function

Function CreateEPG(token As String, tenant As String, anp As String, epg As String, bd As String)
'method:ÊPOST
'url:Êhttp://muc-apic/api/node/mo/uni/tn-testtenant/ap-testanp/epg-testepg.json
'payload:Ê{"fvAEPg":{"attributes":{"dn":"uni/tn-testtenant/ap-testanp/epg-testepg","name":"testepg","rn":"epg-testepg","status":"created"},"children":[{"fvCrtrn":{"attributes":{"dn":"uni/tn-testtenant/ap-testanp/epg-testepg/crtrn","name":"default","rn":"crtrn","status":"created,modified"},"children":[]}},{"fvRsBd":{"attributes":{"tnFvBDName":"default","status":"created,modified"},"children":[]}}]}}
Dim Resource As String, errmsg As String, Body As String
        Resource = "node/mo/uni/tn-" & tenant & "/ap-" & anp & "/epg-" & epg & ".json"
        Body = "{""fvAEPg"":{""attributes"":{""dn"":""uni/tn-" & tenant & "/ap-" & anp & "/epg-" & epg & """,""name"":""" & epg & """," _
        & """rn"":""epg-" & epg & """,""status"":""created""},""children"":[{""fvCrtrn"":{""attributes"":{""dn"":""uni/tn-" & tenant & "/ap-" & anp & "/epg-" & epg & "/crtrn""," _
        & """name"":""default"",""rn"":""crtrn"",""status"":""created,modified""},""children"":[]}},{""fvRsBd"":{""attributes"":{""tnFvBDName"":""" & bd & """,""status"":" _
        & """created,modified""},""children"":[]}}]}}"
        errmsg = "Error creating EPG " & epg & "in ANP " & anp & " in tenant " & tenant
        SendPOST token, Resource, Body, errmsg
End Function

Function CreateVRF(token As String, tenant As String, vrf As String)
'url:Êhttp://muc-apic/api/node/mo/uni/tn-testtenant/ctx-testvrf.json
'payload:Ê{"fvCtx":{"attributes":{"dn":"uni/tn-testtenant/ctx-testvrf","name":"testvrf","rn":"ctx-testvrf","status":"created"},"children":[]}}
Dim Resource As String, errmsg As String, Body As String
        Resource = "node/mo/uni/tn-" & tenant & "/ctx-" & vrf & ".json"
        Body = "{""fvCtx"":{""attributes"":{""dn"":""uni/tn-" & tenant & "/ctx-" & vrf & """,""name"":""" & vrf & """,""rn"":""ctx-" & vrf & """,""status"":""created""},""children"":[]}}"
        errmsg = "Error creating VRF " & vrf & " in tenant " & tenant
        SendPOST token, Resource, Body, errmsg
End Function

Function CreateBD(token As String, tenant As String, vrf As String, bd As String)
'url:Êhttp://muc-apic/api/node/mo/uni/tn-testtenant/BD-testbd.json
'payload:Ê{"fvBD":{"attributes":{"dn":"uni/tn-testtenant/BD-testbd","mac":"00:22:BD:F8:19:FF","name":"testbd","rn":"BD-testbd","status":"created"},"children":[{"fvRsCtx":{"attributes":{"tnFvCtxName":"testvrf","status":"created,modified"},"children":[]}}]}}
Dim Resource As String, errmsg As String, Body As String
        Resource = "node/mo/uni/tn-" & tenant & "/BD-" & bd & ".json"
        Body = "{""fvBD"":{""attributes"":{""dn"":""uni/tn-" & tenant & "/BD-" & bd & """,""mac"":""00:22:BD:F8:19:FF"",""name"":""" & bd & """," _
        & """rn"":""BD-" & bd & """,""status"":""created""},""children"":[{""fvRsCtx"":{""attributes"":{""tnFvCtxName"":""" & vrf & """,""status"":""created,modified""},""children"":[]}}]}}"
        errmsg = "Error creating BD " & bd & " in tenant " & tenant
        SendPOST token, Resource, Body, errmsg
End Function

Function AddSubnet(token As String, tenant As String, bd As String, subnet As String)
'method:ÊPOST
'url:Êhttp://muc-apic/api/node/mo/uni/tn-testtenant/BD-testbd/subnet-[2.2.2.2/24].json
'payload:Ê{"fvSubnet":{"attributes":{"dn":"uni/tn-testtenant/BD-testbd/subnet-[2.2.2.2/24]","ip":"2.2.2.2/24","scope":"public","rn":"subnet-[2.2.2.2/24]","status":"created"},"children":[]}}
Dim Resource As String, errmsg As String, Body As String
        Resource = "node/mo/uni/tn-" & tenant & "/BD-" & bd & "/subnet-\[" & subnet & "\].json"
        Body = "{""fvSubnet"":{""attributes"":{""dn"":""uni/tn-" & tenant & "/BD-" & bd & "/subnet-[" & subnet & "]"",""ip"":""" & subnet & """," _
        & """scope"":""public"",""rn"":""subnet-[" & subnet & "]"",""status"":""created""},""children"":[]}}"
        errmsg = "Error adding subnet " & subnet & " to BD " & bd & " in tenant " & tenant
        SendPOST token, Resource, Body, errmsg
End Function

Function CreateFEXProfile(token As String, fex As String)
'url:Êhttp://muc-apic/api/node/mo/uni/infra/fexprof-103.json
'payload:Ê{"infraFexP":{"attributes":{"dn":"uni/infra/fexprof-103","name":"103","rn":"fexprof-103","status":"created,modified"},"children":[{"infraFexBndlGrp":{"attributes":{"dn":"uni/infra/fexprof-103/fexbundle-103","name":"103","rn":"fexbundle-103","status":"created,modified"},"children":[]}}]}}
Dim Resource As String, errmsg As String, Body As String
        Resource = "node/mo/uni/infra/fexprof-" & fex & ".json"
        Body = "{""infraFexP"":{""attributes"":{""dn"":""uni/infra/fexprof-" & fex & """,""name"":""" & fex & """,""rn"":""fexprof-" & fex & """,""status"":" _
        & """created,modified""},""children"":[{""infraFexBndlGrp"":{""attributes"":{""dn"":""uni/infra/fexprof-" & fex & "/fexbundle-" & fex & """,""name"":""" & fex & """," _
        & """rn"":""fexbundle-" & fex & """,""status"":""created,modified""},""children"":[]}}]}}"
        errmsg = "Error adding FEX interface profile " & fex
        SendPOST token, Resource, Body, errmsg
End Function

Function AddFEXIfSelector(token As String, fex As String, leaf As String, port1 As String, port2 As String)
'url:Êhttp://muc-apic/api/node/mo/uni/infra/accportprof-201/hports-FEX103-typ-range.json
'payload:Ê{"infraHPortS":{"attributes":{"dn":"uni/infra/accportprof-201/hports-FEX103-typ-range","name":"FEX103","rn":"hports-FEX103-typ-range","status":"created,modified"},"children":[{"infraPortBlk":{"attributes":{"dn":"uni/infra/accportprof-201/hports-FEX103-typ-range/portblk-block2","fromPort":"25","toPort":"26","name":"block2","rn":"portblk-block2","status":"created,modified"},"children":[]}},{"infraRsAccBaseGrp":{"attributes":{"tDn":"uni/infra/fexprof-103/fexbundle-103","fexId":"103","status":"created,modified"},"children":[]}}]}}
Dim Resource As String, errmsg As String, Body As String
        Resource = "node/mo/uni/infra/accportprof-" & leaf & "/hports-FEX" & fex & "-typ-range.json"
        Body = "{""infraHPortS"":{""attributes"":{""dn"":""uni/infra/accportprof-" & leaf & "/hports-FEX" & fex & "-typ-range"",""name"":""FEX" & fex & """," _
        & """rn"":""hports-FEX" & fex & "-typ-range"",""status"":""created,modified""},""children"":[{""infraPortBlk"":{""attributes"":{""dn"":" _
        & """uni/infra/accportprof-" & leaf & "/hports-FEX" & fex & "-typ-range/portblk-block2"",""fromPort"":""" & port1 & """,""toPort"":""" & port2 & """," _
        & """name"":""block2"",""rn"":""portblk-block2"",""status"":""created,modified""},""children"":[]}},{""infraRsAccBaseGrp"":{""attributes"":{" _
        & """tDn"":""uni/infra/fexprof-" & fex & "/fexbundle-" & fex & """,""fexId"":""" & fex & """,""status"":""created,modified""},""children"":[]}}]}}"
        errmsg = "Error adding FEX interface selector " & fex & " to leaf " & leaf
        SendPOST token, Resource, Body, errmsg
End Function

Function AddVPC(token As String, switch1 As String, switch2 As String)
'url:Êhttp://muc-apic/api/node/mo/uni/fabric/protpol/expgep-201-202.json
'payload:Ê{"fabricExplicitGEp":{"attributes":{"dn":"uni/fabric/protpol/expgep-201-202","name":"201-202","id":"1","rn":"expgep-201-202","status":"created"},"children":[{"fabricNodePEp":{"attributes":{"dn":"uni/fabric/protpol/expgep-201-202/nodepep-201","id":"201","status":"created","rn":"nodepep-201"},"children":[]}},{"fabricNodePEp":{"attributes":{"dn":"uni/fabric/protpol/expgep-201-202/nodepep-202","id":"202","status":"created","rn":"nodepep-202"},"children":[]}},{"fabricRsVpcInstPol":{"attributes":{"tnVpcInstPolName":"default","status":"created,modified"},"children":[]}}]}}
Dim Resource As String, errmsg As String, Body As String
Dim Name As String, Id As String
        Name = switch1 & "-" & switch2
        Id = Right(switch1, 1)
        Resource = "node/mo/uni/fabric/protpol/expgep-" & switch1 & "-" & switch2 & ".json"
        Body = "{""fabricExplicitGEp"":{""attributes"":{""dn"":""uni/fabric/protpol/expgep-" & Name & """,""name"":""" & Name & """," _
        & """id"":""" & Id & """,""rn"":""expgep-" & Name & """,""status"":""created""},""children"":[{""fabricNodePEp"":{""attributes"":" _
        & "{""dn"":""uni/fabric/protpol/expgep-" & Name & "/nodepep-" & switch1 & """,""id"":""" & switch1 & """,""status"":""created"",""rn"":""nodepep-" & switch1 & """}" _
        & ",""children"":[]}},{""fabricNodePEp"":{""attributes"":{""dn"":""uni/fabric/protpol/expgep-" & Name & "/nodepep-" & switch2 & """,""id"":""" & switch2 & """," _
        & """status"":""created"",""rn"":""nodepep-" & switch2 & """},""children"":[]}},{""fabricRsVpcInstPol"":{""attributes"":{""tnVpcInstPolName"":""default""," _
        & """status"":""created,modified""},""children"":[]}}]}}"
        errmsg = "Error creating VPC between leaves " & switch1 & " and " & switch2
        SendPOST token, Resource, Body, errmsg
End Function

Function TenantExists(token As String, tenant As String) As Boolean
        TenantExists = ObjectExists(token, "node/mo/uni/tn-" & tenant & ".json")
End Function

Function VRFExists(token As String, tenant As String, vrf As String) As Boolean
        VRFExists = ObjectExists(token, "node/mo/uni/tn-" & tenant & "/ctx-" & vrf & ".json")
End Function

Function BDExists(token As String, tenant As String, bd As String) As Boolean
        BDExists = ObjectExists(token, "node/mo/uni/tn-" & tenant & "/BD-" & bd & ".json")
End Function

Function SubnetExists(token As String, tenant As String, bd As String, subnet As String) As Boolean
        SubnetExists = ObjectExists(token, "node/mo/uni/tn-" & tenant & "/BD-" & bd & "/subnet-\[" & subnet & "\].json")
End Function

Function ANPExists(token As String, tenant As String, anp As String) As Boolean
        ANPExists = ObjectExists(token, "node/mo/uni/tn-" & tenant & "/ap-" & anp & ".json")
End Function

Function EPGExists(token As String, tenant As String, anp As String, epg As String) As Boolean
        EPGExists = ObjectExists(token, "node/mo/uni/tn-" & tenant & "/ap-" & anp & "/epg-" & epg & ".json")
End Function

Function CreateLLP(token As String, Name As String, neg As Boolean, speed As String)
'url:Êhttp://muc-apic/api/node/mo/uni/infra/hintfpol-1G_NoNeg.json
'payload:Ê{"fabricHIfPol":{"attributes":{"dn":"uni/infra/hintfpol-1G_NoNeg","name":"1G_NoNeg","speed":"1G","autoNeg":"off","rn":"hintfpol-1G_NoNeg","status":"created"},"children":[]}}
Dim Response As WebResponse, Resource As String, Body As String, errmsg As String
    Resource = "node/mo/uni/infra/hintfpol-" & Name & ".json"
    Body = "{""fabricHIfPol"":{""attributes"":{""dn"":""uni/infra/hintfpol-" & Name & """,""name"":""" & Name & """,""speed"":""" & speed & """," _
    & """autoNeg"":""" & IIf(neg, "on", "off") & """,""rn"":""hintfpol-" & Name & """,""status"":""created""},""children"":[]}}"
    errmsg = "Error creating link-level policy " & Name
    SendPOST token, Resource, Body, errmsg
End Function

Function CreateCDPP(token As String, Name As String, enabled As Boolean)
'url: http://muc-apic/api/node/mo/uni/infra/cdpIfP-TEST.json
'payload: {"cdpIfPol":{"attributes":{"dn":"uni/infra/cdpIfP-TEST","name":"TEST","rn":"cdpIfP-TEST","status":"created"},"children":[]}}
Dim Response As WebResponse, Resource As String, Body As String, errmsg As String
    Resource = "node/mo/uni/infra/cdpIfP-" & Name & ".json"
    Body = "{""cdpIfPol"":{""attributes"":{""dn"":""uni/infra/cdpIfP-" & Name & """,""name"":""" & Name & """,""adminSt"":""" & IIf(enabled, "enabled", "disabled") & """," _
    & """rn"":""cdpIfP-" & Name & """,""status"":""created""},""children"":[]}}"
    errmsg = "Error creating CDP policy " & Name
    SendPOST token, Resource, Body, errmsg
End Function

Function CreateLLDPP(token As String, Name As String, enabled As Boolean)
Dim Response As WebResponse, Resource As String, Body As String, errmsg As String
    Resource = "node/mo/uni/infra/lldpIfP-" & Name & ".json"
    Body = "{""lldpIfPol"":{""attributes"":{""dn"":""uni/infra/lldpIfP-" & Name & """,""name"":""" & Name & """,""adminRxSt"":""" & IIf(enabled, "enabled", "disabled") & """," _
    & """adminTxSt"":""" & IIf(enabled, "enabled", "disabled") & """rn"":""lldpIfP-" & Name & """,""status"":""created""},""children"":[]}}"
    errmsg = "Error creating LLDP policy " & Name
    SendPOST token, Resource, Body, errmsg
End Function

Function CreateIntPolGroup(token As String, Name As String, llp As String, cdp As String, lldp As String, aep As String)
'url: http://muc-apic/api/node/mo/uni/infra/funcprof/accportgrp-TEST.json
'payload: {"infraAccPortGrp":{"attributes":{"dn":"uni/infra/funcprof/accportgrp-TEST","name":"TEST","rn":"accportgrp-TEST","status":"created"},"children":[{"infraRsAttEntP":{"attributes":{"tDn":"uni/infra/attentp-default","status":"created,modified"},"children":[]}},{"infraRsHIfPol":{"attributes":{"tnFabricHIfPolName":"100M_Neg","status":"created,modified"},"children":[]}},{"infraRsCdpIfPol":{"attributes":{"tnCdpIfPolName":"default","status":"created,modified"},"children":[]}},{"infraRsLldpIfPol":{"attributes":{"tnLldpIfPolName":"default","status":"created,modified"},"children":[]}}]}}
Dim Response As WebResponse, Resource As String, Body As String, errmsg As String
    Resource = "node/mo/uni/infra/funcprof/accportgrp-" & Name & ".json"
    Body = "{""infraAccPortGrp"":{""attributes"":{""dn"":""uni/infra/funcprof/accportgrp-" & Name & """,""name"":""" & Name & """,""rn"":""accportgrp-" & Name & """," _
    & """status"":""created""},""children"":[{""infraRsAttEntP"":{""attributes"":{""tDn"":""uni/infra/attentp-" & aep & """,""status"":""created,modified""}," _
    & """children"":[]}},{""infraRsHIfPol"":{""attributes"":{""tnFabricHIfPolName"":""" & llp & """,""status"":""created,modified""},""children"":[]}}," _
    & "{""infraRsCdpIfPol"":{""attributes"":{""tnCdpIfPolName"":""" & cdp & """,""status"":""created,modified""},""children"":[]}}," _
    & "{""infraRsLldpIfPol"":{""attributes"":{""tnLldpIfPolName"":""" & lldp & """,""status"":""created,modified""},""children"":[]}}]}}"
    errmsg = "Error creating policy group " & Name
    SendPOST token, Resource, Body, errmsg
End Function

Function CreateAAEP(token As String, Name As String)
'url: http://muc-apic/api/node/mo/uni/infra.json
'payload: {"infraInfra":{"attributes":{"dn":"uni/infra","status":"modified"},"children":[{"infraAttEntityP":{"attributes":{"dn":"uni/infra/attentp-Zone1","name":"Zone1","rn":"attentp-Zone1","status":"created"},"children":[]}},{"infraFuncP":{"attributes":{"dn":"uni/infra/funcprof","status":"modified"},"children":[]}}]}}
Dim Response As WebResponse, Resource As String, Body As String, errmsg As String
    Resource = "node/mo/uni/infra.json"
    Body = "{""infraInfra"":{""attributes"":{""dn"":""uni/infra"",""status"":""modified""},""children"":[{""infraAttEntityP"":{""attributes"":{" _
    & """dn"":""uni/infra/attentp-" & Name & """,""name"":""" & Name & """,""rn"":""attentp-" & Name & """,""status"":""created""},""children"":[]}}," _
    & "{""infraFuncP"":{""attributes"":{""dn"":""uni/infra/funcprof"",""status"":""modified""},""children"":[]}}]}}"
    errmsg = "Error creating policy group " & Name
    SendPOST token, Resource, Body, errmsg
End Function

Function CreateStaticVLANPool(token As String, Name As String, from_vlan As String, to_vlan As String)
'url: http://muc-apic/api/node/mo/uni/infra/vlanns-[phys]-static.json
'payload: {"fvnsVlanInstP":{"attributes":{"dn":"uni/infra/vlanns-[phys]-static","name":"phys","allocMode":"static","rn":"vlanns-[phys]-static","status":"created"},"children":[{"fvnsEncapBlk":{"attributes":{"dn":"uni/infra/vlanns-[phys]-static/from-[vlan-1]-to-[vlan-99]","from":"vlan-1","to":"vlan-99","rn":"from-[vlan-1]-to-[vlan-99]","status":"created"},"children":[]}}]}}
Dim Response As WebResponse, Resource As String, Body As String, errmsg As String
    Resource = "node/mo/uni/infra/vlanns-\[" & Name & "\]-static.json"
    Body = "{""fvnsVlanInstP"":{""attributes"":{""dn"":""uni/infra/vlanns-[" & Name & "]-static"",""name"":""" & Name & """,""allocMode"":""static""," _
    & """rn"":""vlanns-[" & Name & "]-static"",""status"":""created""},""children"":[{""fvnsEncapBlk"":{""attributes"":" _
    & "{""dn"":""uni/infra/vlanns-[" & Name & "]-static/from-[vlan-" & from_vlan & "]-to-[vlan-" & to_vlan & "]""," _
    & """from"":""vlan-" & from_vlan & """,""to"":""vlan-" & to_vlan & """,""rn"":""from-[vlan-" & from_vlan & "]-to-[vlan-" & to_vlan & "]"",""status"":""created""},""children"":[]}}]}}"
    errmsg = "Error creating static VLAN pool " & Name
    SendPOST token, Resource, Body, errmsg
End Function

Function CreateDynamicVLANPool(token As String, Name As String, from_vlan As String, to_vlan As String)
'url: http://muc-apic/api/node/mo/uni/infra/vlanns-[dyntest]-dynamic.json
'payload{"fvnsVlanInstP":{"attributes":{"dn":"uni/infra/vlanns-[dyntest]-dynamic","name":"dyntest","rn":"vlanns-[dyntest]-dynamic","status":"created"},
'  "children":[{"fvnsEncapBlk":{"attributes":{"dn":"uni/infra/vlanns-[dyntest]-dynamic/from-[vlan-501]-to-[vlan-502]",
'  "from":"vlan-501","to":"vlan-502","rn":"from-[vlan-501]-to-[vlan-502]","status":"created"},"children":[]}}]}}
Dim Response As WebResponse, Resource As String, Body As String, errmsg As String
    Resource = "node/mo/uni/infra/vlanns-\[" & Name & "\]-dynamic.json"
    Body = "{""fvnsVlanInstP"":{""attributes"":{""dn"":""uni/infra/vlanns-[" & Name & "]-dynamic"",""name"":""" & Name & """,""allocMode"":""dynamic""," _
    & """rn"":""vlanns-[" & Name & "]-dynamic"",""status"":""created,modified""},""children"":[{""fvnsEncapBlk"":{""attributes"":" _
    & "{""dn"":""uni/infra/vlanns-[" & Name & "]-dynamic/from-[vlan-" & from_vlan & "]-to-[vlan-" & to_vlan & "]""," _
    & """from"":""vlan-" & from_vlan & """,""to"":""vlan-" & to_vlan & """,""rn"":""from-[vlan-" & from_vlan & "]-to-[vlan-" & to_vlan & "]"",""status"":""created,modified""},""children"":[]}}]}}"
    errmsg = "Error creating dynamic VLAN pool " & Name
    SendPOST token, Resource, Body, errmsg
End Function

Function CreatePhysDomain(token As String, Name As String, pool As String)
'url: http://muc-apic/api/node/mo/uni/phys-Zone1.json
'payload: {"physDomP":{"attributes":{"dn":"uni/phys-Zone1","name":"Zone1","rn":"phys-Zone1","status":"created"},"children":[{"infraRsVlanNs":{"attributes":{"tDn":"uni/infra/vlanns-[Zone1]-static","status":"created"},"children":[]}}]}}
Dim Response As WebResponse, Resource As String, Body As String, errmsg As String
    Resource = "node/mo/uni/phys-" & Name & ".json"
    Body = "{""physDomP"":{""attributes"":{""dn"":""uni/phys-" & Name & """,""name"":""" & Name & """,""rn"":""phys-" & Name & """,""status"":""created""}," _
    & """children"":[{""infraRsVlanNs"":{""attributes"":{""tDn"":""uni/infra/vlanns-[" & pool & "]-static"",""status"":""created""},""children"":[]}}]}}"
    errmsg = "Error creating physical domain " & Name
    SendPOST token, Resource, Body, errmsg
End Function

Function AddPhysDomainToAEP(token As String, domain As String, aep As String)
'url: http://muc-apic/api/node/mo/uni/infra/attentp-Zone1.json
'payload: {"infraRsDomP":{"attributes":{"tDn":"uni/phys-Zone2","status":"created,modified"},"children":[]}}
Dim Response As WebResponse, Resource As String, Body As String, errmsg As String
    Resource = "node/mo/uni/infra/attentp-" & aep & ".json"
    Body = "{""infraRsDomP"":{""attributes"":{""tDn"":""uni/phys-" & domain & """,""status"":""created,modified""},""children"":[]}}"
    errmsg = "Error assigning physical domain " & domain & " to AEP " & aep
    SendPOST token, Resource, Body, errmsg
End Function

Function AddVMMDomainToAEP(token As String, domain As String, aep As String)
'url: http://muc-apic/api/node/mo/uni/infra/attentp-default.json
'payload{"infraRsDomP":{"attributes":{"tDn":"uni/vmmp-VMware/dom-testJose","status":"created"},"children":[]}}
Dim Response As WebResponse, Resource As String, Body As String, errmsg As String
    Resource = "node/mo/uni/infra/attentp-" & aep & ".json"
    Body = "{""infraRsDomP"":{""attributes"":{""tDn"":""uni/vmmp-VMware/dom-" & domain & """,""status"":""created,modified""},""children"":[]}}"
    errmsg = "Error assigning virtual domain " & domain & " to AEP " & aep
    SendPOST token, Resource, Body, errmsg
End Function

Function AddPhysDomainToEPG(token As String, domain As String, tenant As String, anp As String, epg As String)
'url: http://muc-apic/api/node/mo/uni/tn-common/ap-default/epg-VLAN1.json
'payload: {"fvRsDomAtt":{"attributes":{"instrImedcy":"immediate","resImedcy":"immediate","tDn":"uni/phys-default","status":"created"},"children":[]}}
Dim Response As WebResponse, Resource As String, Body As String, errmsg As String
    Resource = "node/mo/uni/tn-" & tenant & "/ap-" & anp & "/epg-" & epg & ".json"
    Body = "{""fvRsDomAtt"":{""attributes"":{""instrImedcy"":""immediate"",""resImedcy"":""immediate"",""tDn"":""uni/phys-" & domain & """,""status"":""created,modified""},""children"":[]}}"
    errmsg = "Error assigning physical domain " & domain & " to EPG " & tenant & "/" & anp & "/" & epg
    SendPOST token, Resource, Body, errmsg
End Function

Function AddVirtDomainToEPG(token As String, domain As String, tenant As String, anp As String, epg As String)
'url: http://muc-apic/api/node/mo/uni/tn-Pod2/ap-Pod2/epg-EPG2.json
'payload{"fvRsDomAtt":{"attributes":{"resImedcy":"immediate","instrImedcy":"immediate","tDn":"uni/vmmp-VMware/dom-ACI-vCenter-VDS","status":"created"},"children":[{"vmmSecP":{"attributes":{"status":"created"},"children":[]}}]}}
Dim Response As WebResponse, Resource As String, Body As String, errmsg As String
    Resource = "node/mo/uni/tn-" & tenant & "/ap-" & anp & "/epg-" & epg & ".json"
    Body = "{""fvRsDomAtt"":{""attributes"":{""instrImedcy"":""immediate"",""resImedcy"":""immediate"",""tDn"":""uni/vmmp-VMware/dom-" & domain & """,""status"":""created,modified""},""children"":[]}}"
    errmsg = "Error assigning physical domain " & domain & " to EPG " & tenant & "/" & anp & "/" & epg
    SendPOST token, Resource, Body, errmsg
End Function

Function DeleteStaticBindingFEXVPC(token As String, tenant As String, anp As String, epg As String, intprofile, fexparent, servername)
'url: http://muc-apic/api/node/mo/uni/tn-Acme/ap-MyApp1/epg-Tier1.json
'payload: {"fvAEPg":{"attributes":{"dn":"uni/tn-common/ap-default/epg-VLAN2","status":"modified"},
'  "children":[{"fvRsPathAtt":{"attributes":{"dn":"uni/tn-common/ap-default/epg-VLAN2/rspathAtt-[topology/pod-1/protpaths-201-202/extprotpaths-101-102/pathep-[OOB-L2]]","status":"deleted"},"children":[]}}]}}
Dim Resource As String, errmsg As String, Body As String
        Resource = "node/mo/uni/tn-" & tenant & "/ap-" & anp & "/epg-" & epg & ".json"
        Body = "{""fvAEPg"":{""attributes"":{""dn"":""uni/tn-" & tenant & "/ap-" & anp & "/epg-" & epg & """,""status"":""modified""}," _
        & """children"":[{""fvRsPathAtt"":{""attributes"":{""dn"":""uni/tn-" & tenant & "/ap-" & anp & "/epg-" & epg & "/rspathAtt-" _
        & "[topology/pod-1/protpaths-" & fexparent & "/extprotpaths-" & intprofile & "/pathep-[" & servername & "]]"",""status"":""deleted""},""children"":[]}}]}}"
        errmsg = "Error in static binding delete request"
        SendPOST token, Resource, Body, errmsg
End Function

Function GetSwitches(token As String, url) As String()
Dim i, aux() As String
        'Get the list of nodes
        Dim ACIclient As New WebClient
        ACIclient.BaseUrl = url
        'Create Request
        Dim ACIrequest As New WebRequest
        ACIrequest.Resource = "node/class/fabricNode.json"
        ACIrequest.method = WebMethod.HttpGet
        ACIrequest.Body = "http://muc-apic/api/node/class/fabricNode.json}"
        ACIrequest.SetHeader "Cookie", "APIC-cookie=" & token
        ACIrequest.SetHeader "Content-Type", "application/json"
        'Send request and store response
        Dim Response As WebResponse
        Set Response = ACIclient.Execute(ACIrequest)
        If Not Response.StatusCode = 200 Then
            MsgBox "Error getting the node list: " & Response.StatusCode & " - " & Response.StatusDescription
            Exit Function
        End If
        'Process response
        Dim Json As Object
        Set Json = JsonConverter.ParseJson(Response.Content)
        For i = 1 To Json("totalCount")
            If Json("imdata").Item(i)("fabricNode")("attributes")("role") = "leaf" Then
                'MsgBox "Leaf " & Json("imdata").Item(i)("fabricNode")("attributes")("id") & " detected"
                ReDim Preserve aux(i)
                aux(i) = Json("imdata").Item(i)("fabricNode")("attributes")("id")
            End If
        Next i
        GetSwitches = aux
End Function












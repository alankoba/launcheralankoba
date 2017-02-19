VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private m_CRC As Long
Private m_CRCAsm() As Byte
Private m_CRCTable(0 To 255) As Long

Function CRC(FileName As String) As Variant
Dim Filenr As Integer
Dim ByteArray() As Byte
Dim ByteArray2() As Byte
Dim n(2) As Long
On Error Resume Next
If (FileLen(FileName) = 0) Then
Exit Function
End If
Call InitializeCRC
Filenr = FreeFile
Open FileName For Binary As #Filenr
n(2) = LOF(Filenr)
n(0) = 1
n(1) = 1048576
m_CRC = -1
Volta:
If CLng((n(0) + n(1)) - 1) > n(2) Then
    n(1) = (n(2) - n(0)) + 1
    ReDim ByteArray2(0 To n(1) - 1)
    Get #Filenr, n(0), ByteArray2()
    CRC = CRC + AddBytes(ByteArray2, n(1))
    Close #Filenr
    Exit Function
End If
  ReDim ByteArray(0 To n(1) - 1)
  Get #Filenr, n(0), ByteArray()
  CRC = CRC + AddBytes(ByteArray, n(1))
n(0) = n(0) + n(1)
GoTo Volta:
End Function

Private Function AddBytes(ByteArray() As Byte, ByteSize As Long) As Variant
On Local Error GoTo NoData
Call CallWindowProc(VarPtr(m_CRCAsm(0)), VarPtr(m_CRC), VarPtr(ByteArray(LBound(ByteArray))), VarPtr(m_CRCTable(0)), ByteSize)
NoData:
AddBytes = (Not m_CRC)
End Function

Private Sub InitializeCRC()
Dim I As Long
Dim sASM As String
m_CRCTable(0) = &H0
m_CRCTable(1) = &H77073096
m_CRCTable(2) = &HEE0E612C
m_CRCTable(3) = &H990951BA
m_CRCTable(4) = &H76DC419
m_CRCTable(5) = &H706AF48F
m_CRCTable(6) = &HE963A535
m_CRCTable(7) = &H9E6495A3
m_CRCTable(8) = &HEDB8832
m_CRCTable(9) = &H79DCB8A4
m_CRCTable(10) = &HE0D5E91E
m_CRCTable(11) = &H97D2D988
m_CRCTable(12) = &H9B64C2B
m_CRCTable(13) = &H7EB17CBD
m_CRCTable(14) = &HE7B82D07
m_CRCTable(15) = &H90BF1D91
m_CRCTable(16) = &H1DB71064
m_CRCTable(17) = &H6AB020F2
m_CRCTable(18) = &HF3B97148
m_CRCTable(19) = &H84BE41DE
m_CRCTable(20) = &H1ADAD47D
m_CRCTable(21) = &H6DDDE4EB
m_CRCTable(22) = &HF4D4B551
m_CRCTable(23) = &H83D385C7
m_CRCTable(24) = &H136C9856
m_CRCTable(25) = &H646BA8C0
m_CRCTable(26) = &HFD62F97A
m_CRCTable(27) = &H8A65C9EC
m_CRCTable(28) = &H14015C4F
m_CRCTable(29) = &H63066CD9
m_CRCTable(30) = &HFA0F3D63
m_CRCTable(31) = &H8D080DF5
m_CRCTable(32) = &H3B6E20C8
m_CRCTable(33) = &H4C69105E
m_CRCTable(34) = &HD56041E4
m_CRCTable(35) = &HA2677172
m_CRCTable(36) = &H3C03E4D1
m_CRCTable(37) = &H4B04D447
m_CRCTable(38) = &HD20D85FD
m_CRCTable(39) = &HA50AB56B
m_CRCTable(40) = &H35B5A8FA
m_CRCTable(41) = &H42B2986C
m_CRCTable(42) = &HDBBBC9D6
m_CRCTable(43) = &HACBCF940
m_CRCTable(44) = &H32D86CE3
m_CRCTable(45) = &H45DF5C75
m_CRCTable(46) = &HDCD60DCF
m_CRCTable(47) = &HABD13D59
m_CRCTable(48) = &H26D930AC
m_CRCTable(49) = &H51DE003A
m_CRCTable(50) = &HC8D75180
m_CRCTable(51) = &HBFD06116
m_CRCTable(52) = &H21B4F4B5
m_CRCTable(53) = &H56B3C423
m_CRCTable(54) = &HCFBA9599
m_CRCTable(55) = &HB8BDA50F
m_CRCTable(56) = &H2802B89E
m_CRCTable(57) = &H5F058808
m_CRCTable(58) = &HC60CD9B2
m_CRCTable(59) = &HB10BE924
m_CRCTable(60) = &H2F6F7C87
m_CRCTable(61) = &H58684C11
m_CRCTable(62) = &HC1611DAB
m_CRCTable(63) = &HB6662D3D
m_CRCTable(64) = &H76DC4190
m_CRCTable(65) = &H1DB7106
m_CRCTable(66) = &H98D220BC
m_CRCTable(67) = &HEFD5102A
m_CRCTable(68) = &H71B18589
m_CRCTable(69) = &H6B6B51F
m_CRCTable(70) = &H9FBFE4A5
m_CRCTable(71) = &HE8B8D433
m_CRCTable(72) = &H7807C9A2
m_CRCTable(73) = &HF00F934
m_CRCTable(74) = &H9609A88E
m_CRCTable(75) = &HE10E9818
m_CRCTable(76) = &H7F6A0DBB
m_CRCTable(77) = &H86D3D2D
m_CRCTable(78) = &H91646C97
m_CRCTable(79) = &HE6635C01
m_CRCTable(80) = &H6B6B51F4
m_CRCTable(81) = &H1C6C6162
m_CRCTable(82) = &H856530D8
m_CRCTable(83) = &HF262004E
m_CRCTable(84) = &H6C0695ED
m_CRCTable(85) = &H1B01A57B
m_CRCTable(86) = &H8208F4C1
m_CRCTable(87) = &HF50FC457
m_CRCTable(88) = &H65B0D9C6
m_CRCTable(89) = &H12B7E950
m_CRCTable(90) = &H8BBEB8EA
m_CRCTable(91) = &HFCB9887C
m_CRCTable(92) = &H62DD1DDF
m_CRCTable(93) = &H15DA2D49
m_CRCTable(94) = &H8CD37CF3
m_CRCTable(95) = &HFBD44C65
m_CRCTable(96) = &H4DB26158
m_CRCTable(97) = &H3AB551CE
m_CRCTable(98) = &HA3BC0074
m_CRCTable(99) = &HD4BB30E2
m_CRCTable(100) = &H4ADFA541
m_CRCTable(101) = &H3DD895D7
m_CRCTable(102) = &HA4D1C46D
m_CRCTable(103) = &HD3D6F4FB
m_CRCTable(104) = &H4369E96A
m_CRCTable(105) = &H346ED9FC
m_CRCTable(106) = &HAD678846
m_CRCTable(107) = &HDA60B8D0
m_CRCTable(108) = &H44042D73
m_CRCTable(109) = &H33031DE5
m_CRCTable(110) = &HAA0A4C5F
m_CRCTable(111) = &HDD0D7CC9
m_CRCTable(112) = &H5005713C
m_CRCTable(113) = &H270241AA
m_CRCTable(114) = &HBE0B1010
m_CRCTable(115) = &HC90C2086
m_CRCTable(116) = &H5768B525
m_CRCTable(117) = &H206F85B3
m_CRCTable(118) = &HB966D409
m_CRCTable(119) = &HCE61E49F
m_CRCTable(120) = &H5EDEF90E
m_CRCTable(121) = &H29D9C998
m_CRCTable(122) = &HB0D09822
m_CRCTable(123) = &HC7D7A8B4
m_CRCTable(124) = &H59B33D17
m_CRCTable(125) = &H2EB40D81
m_CRCTable(126) = &HB7BD5C3B
m_CRCTable(127) = &HC0BA6CAD
m_CRCTable(128) = &HEDB88320
m_CRCTable(129) = &H9ABFB3B6
m_CRCTable(130) = &H3B6E20C
m_CRCTable(131) = &H74B1D29A
m_CRCTable(132) = &HEAD54739
m_CRCTable(133) = &H9DD277AF
m_CRCTable(134) = &H4DB2615
m_CRCTable(135) = &H73DC1683
m_CRCTable(136) = &HE3630B12
m_CRCTable(137) = &H94643B84
m_CRCTable(138) = &HD6D6A3E
m_CRCTable(139) = &H7A6A5AA8
m_CRCTable(140) = &HE40ECF0B
m_CRCTable(141) = &H9309FF9D
m_CRCTable(142) = &HA00AE27
m_CRCTable(143) = &H7D079EB1
m_CRCTable(144) = &HF00F9344
m_CRCTable(145) = &H8708A3D2
m_CRCTable(146) = &H1E01F268
m_CRCTable(147) = &H6906C2FE
m_CRCTable(148) = &HF762575D
m_CRCTable(149) = &H806567CB
m_CRCTable(150) = &H196C3671
m_CRCTable(151) = &H6E6B06E7
m_CRCTable(152) = &HFED41B76
m_CRCTable(153) = &H89D32BE0
m_CRCTable(154) = &H10DA7A5A
m_CRCTable(155) = &H67DD4ACC
m_CRCTable(156) = &HF9B9DF6F
m_CRCTable(157) = &H8EBEEFF9
m_CRCTable(158) = &H17B7BE43
m_CRCTable(159) = &H60B08ED5
m_CRCTable(160) = &HD6D6A3E8
m_CRCTable(161) = &HA1D1937E
m_CRCTable(162) = &H38D8C2C4
m_CRCTable(163) = &H4FDFF252
m_CRCTable(164) = &HD1BB67F1
m_CRCTable(165) = &HA6BC5767
m_CRCTable(166) = &H3FB506DD
m_CRCTable(167) = &H48B2364B
m_CRCTable(168) = &HD80D2BDA
m_CRCTable(169) = &HAF0A1B4C
m_CRCTable(170) = &H36034AF6
m_CRCTable(171) = &H41047A60
m_CRCTable(172) = &HDF60EFC3
m_CRCTable(173) = &HA867DF55
m_CRCTable(174) = &H316E8EEF
m_CRCTable(175) = &H4669BE79
m_CRCTable(176) = &HCB61B38C
m_CRCTable(177) = &HBC66831A
m_CRCTable(178) = &H256FD2A0
m_CRCTable(179) = &H5268E236
m_CRCTable(180) = &HCC0C7795
m_CRCTable(181) = &HBB0B4703
m_CRCTable(182) = &H220216B9
m_CRCTable(183) = &H5505262F
m_CRCTable(184) = &HC5BA3BBE
m_CRCTable(185) = &HB2BD0B28
m_CRCTable(186) = &H2BB45A92
m_CRCTable(187) = &H5CB36A04
m_CRCTable(188) = &HC2D7FFA7
m_CRCTable(189) = &HB5D0CF31
m_CRCTable(190) = &H2CD99E8B
m_CRCTable(191) = &H5BDEAE1D
m_CRCTable(192) = &H9B64C2B0
m_CRCTable(193) = &HEC63F226
m_CRCTable(194) = &H756AA39C
m_CRCTable(195) = &H26D930A
m_CRCTable(196) = &H9C0906A9
m_CRCTable(197) = &HEB0E363F
m_CRCTable(198) = &H72076785
m_CRCTable(199) = &H5005713
m_CRCTable(200) = &H95BF4A82
m_CRCTable(201) = &HE2B87A14
m_CRCTable(202) = &H7BB12BAE
m_CRCTable(203) = &HCB61B38
m_CRCTable(204) = &H92D28E9B
m_CRCTable(205) = &HE5D5BE0D
m_CRCTable(206) = &H7CDCEFB7
m_CRCTable(207) = &HBDBDF21
m_CRCTable(208) = &H86D3D2D4
m_CRCTable(209) = &HF1D4E242
m_CRCTable(210) = &H68DDB3F8
m_CRCTable(211) = &H1FDA836E
m_CRCTable(212) = &H81BE16CD
m_CRCTable(213) = &HF6B9265B
m_CRCTable(214) = &H6FB077E1
m_CRCTable(215) = &H18B74777
m_CRCTable(216) = &H88085AE6
m_CRCTable(217) = &HFF0F6A70
m_CRCTable(218) = &H66063BCA
m_CRCTable(219) = &H11010B5C
m_CRCTable(220) = &H8F659EFF
m_CRCTable(221) = &HF862AE69
m_CRCTable(222) = &H616BFFD3
m_CRCTable(223) = &H166CCF45
m_CRCTable(224) = &HA00AE278
m_CRCTable(225) = &HD70DD2EE
m_CRCTable(226) = &H4E048354
m_CRCTable(227) = &H3903B3C2
m_CRCTable(228) = &HA7672661
m_CRCTable(229) = &HD06016F7
m_CRCTable(230) = &H4969474D
m_CRCTable(231) = &H3E6E77DB
m_CRCTable(232) = &HAED16A4A
m_CRCTable(233) = &HD9D65ADC
m_CRCTable(234) = &H40DF0B66
m_CRCTable(235) = &H37D83BF0
m_CRCTable(236) = &HA9BCAE53
m_CRCTable(237) = &HDEBB9EC5
m_CRCTable(238) = &H47B2CF7F
m_CRCTable(239) = &H30B5FFE9
m_CRCTable(240) = &HBDBDF21C
m_CRCTable(241) = &HCABAC28A
m_CRCTable(242) = &H53B39330
m_CRCTable(243) = &H24B4A3A6
m_CRCTable(244) = &HBAD03605
m_CRCTable(245) = &HCDD70693
m_CRCTable(246) = &H54DE5729
m_CRCTable(247) = &H23D967BF
m_CRCTable(248) = &HB3667A2E
m_CRCTable(249) = &HC4614AB8
m_CRCTable(250) = &H5D681B02
m_CRCTable(251) = &H2A6F2B94
m_CRCTable(252) = &HB40BBE37
m_CRCTable(253) = &HC30C8EA1
m_CRCTable(254) = &H5A05DF1B
m_CRCTable(255) = &H2D02EF8D
sASM = "5589E557565053518B45088B008B750C8B7D108B4D1431DB8A1E30C3C1E80833049F464975F28B4D088901595B585E5F89EC5DC21000"
ReDim m_CRCAsm(0 To Len(sASM) \ 2 - 1)
For I = 1 To Len(sASM) Step 2
  m_CRCAsm(I \ 2) = Val("&H" & Mid$(sASM, I, 2))
Next
End Sub



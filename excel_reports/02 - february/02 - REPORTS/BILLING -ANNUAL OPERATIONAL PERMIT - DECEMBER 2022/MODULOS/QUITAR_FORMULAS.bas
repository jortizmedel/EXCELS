Attribute VB_Name = "Módulo2"
Sub QuitarFormulas()
    
    Range("A2").Select 'Name
    Selection.CurrentRegion.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range("D2").Select 'Address
    Selection.CurrentRegion.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range("F2").Select 'City
    Selection.CurrentRegion.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("G2").Select 'State
    Selection.CurrentRegion.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range("H2").Select 'ZipCode
    Selection.CurrentRegion.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range("I2").Select 'PhoneEmail
    Selection.CurrentRegion.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range("K2").Select 'MailingAddress
    Selection.CurrentRegion.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("L2").Select 'MailingCity
    Selection.CurrentRegion.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range("N2").Select 'MailingState
    Selection.CurrentRegion.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("O2").Select 'MailingZip
    Selection.CurrentRegion.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("P2").Select 'Extra1
    Selection.CurrentRegion.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("R2").Select 'Extra2
    Selection.CurrentRegion.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("V2").Select 'AddField2
    Selection.CurrentRegion.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range("Z2").Select 'NOTE
    Selection.CurrentRegion.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range("AB2").Select 'PermitFee
    Selection.CurrentRegion.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Sub EliminarFilasVacias()
    Rows("1:1").Select
    Selection.AutoFilter
    Range("D5").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveSheet.Range("$A$1:$AB$893").AutoFilter Field:=28, Criteria1:="="
    Selection.SpecialCells(xlCellTypeVisible).Select
    Union(Range( _
        "91:91,93:93,96:96,99:99,101:101,103:103,105:105,107:107,109:109,111:111,114:114,116:116,118:118,120:120,123:123,126:126,129:129,131:131,134:134,137:137,139:139,141:141,143:143,145:145,147:147,149:149,151:151,154:154,156:156,158:158,160:161,163:163" _
        ), Range( _
        "165:165,167:167,169:169,172:172,174:175,177:177,181:182,184:184,186:186,188:188,193:193,195:195,198:198,201:201,204:204,210:210,213:213,215:215,217:217,219:219,222:222,224:224,226:226,228:228,232:232,234:234,236:238,240:240,242:262,264:284,286:306,308:308" _
        ), Range( _
        "311:311,314:314,316:316,318:318,320:320,322:322,325:325,328:329,331:332,334:334,337:337,339:339,344:344,346:346,349:349,351:351,355:355,357:357,359:359,361:361,369:369,373:373,376:376,379:380,382:382,384:384,390:390,393:393,395:395,397:397,399:399,401:401" _
        ), Range( _
        "404:404,408:408,410:410,412:412,415:415,418:418,420:420,422:424,427:427,431:431,434:434,437:437,442:442,444:444,446:447,450:450,457:457,460:460,463:463,465:465,469:470,472:472,476:478,481:481,485:486,488:489,491:492,494:495,497:498,503:503,505:506,509:509" _
        ), Range( _
        "518:518,520:521,525:526,528:528,530:531,533:533,537:538,541:542,545:545,547:547,550:550,552:552,554:554,573:574,576:576,578:578,580:580,582:582,584:584,589:590,592:592,595:595,601:601,604:604,609:609,613:613,619:619,621:621,624:624,626:626,633:633,635:636" _
        ), Range( _
        "638:638,642:642,647:647,652:652,654:654,656:656,658:658,661:661,663:663,666:666,668:668,670:670,672:672,674:675,677:693,695:695,697:697,702:702,704:704,708:709,722:722,724:724,726:728,730:731,733:734,736:737,741:742,747:747,751:752,755:756,762:762,767:771" _
        ), Range( _
        "775:775,778:779,782:782,786:789,792:792,795:795,802:802,805:805,811:811,816:816,820:821,823:1048576,5:5,7:7,9:9,12:12,16:16,19:20,22:22,24:24,26:26,28:31,33:33,35:35,37:37,39:39,41:41,43:43,45:45,47:47,49:50,52:52" _
        ), Range( _
        "54:57,61:61,63:63,65:65,67:67,70:70,73:74,76:76,80:80,83:83,85:85,88:88")). _
        Select
    Range("A823").Activate
    Selection.Delete Shift:=xlUp
    ActiveWindow.ScrollRow = 809
    ActiveWindow.ScrollRow = 805
    ActiveWindow.ScrollRow = 801
    ActiveWindow.ScrollRow = 792
    ActiveWindow.ScrollRow = 782
    ActiveWindow.ScrollRow = 773
    ActiveWindow.ScrollRow = 763
    ActiveWindow.ScrollRow = 751
    ActiveWindow.ScrollRow = 734
    ActiveWindow.ScrollRow = 715
    ActiveWindow.ScrollRow = 700
    ActiveWindow.ScrollRow = 682
    ActiveWindow.ScrollRow = 663
    ActiveWindow.ScrollRow = 646
    ActiveWindow.ScrollRow = 629
    ActiveWindow.ScrollRow = 613
    ActiveWindow.ScrollRow = 598
    ActiveWindow.ScrollRow = 583
    ActiveWindow.ScrollRow = 569
    ActiveWindow.ScrollRow = 560
    ActiveWindow.ScrollRow = 550
    ActiveWindow.ScrollRow = 537
    ActiveWindow.ScrollRow = 529
    ActiveWindow.ScrollRow = 521
    ActiveWindow.ScrollRow = 508
    ActiveWindow.ScrollRow = 502
    ActiveWindow.ScrollRow = 485
    ActiveWindow.ScrollRow = 479
    ActiveWindow.ScrollRow = 471
    ActiveWindow.ScrollRow = 466
    ActiveWindow.ScrollRow = 458
    ActiveWindow.ScrollRow = 381
    ActiveWindow.LargeScroll Down:=-1
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveSheet.Range("$A$1:$AB$455").AutoFilter Field:=28
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Rows("1:1").Select
    Selection.AutoFilter
End Sub



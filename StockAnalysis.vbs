Sub StockAnalysis()

    YearSet = Array("2014", "2015", "2016")
    
    ' Iterate through each Year

    For Each DataTab In YearSet
    
        MsgBox ("Processing Year " + DataTab)
    
        EndOfData = False
        DataRow = 2
        OutputRow = 2
        
        Ticker = ""
        TickerSumVolume = 0
        PreviousTicker = ""
        
        StartPrice = 0
        EndPrice = 0
        YearlyChange = 0
        PercentageChange = 0
        
        TopPerformerTicker = ""
        TopPerformerValue = 0
        BottomPerformerTicker = ""
        BottomPerformerValue = 0
        HighestVolumeTicker = ""
        HighestVolumeValue = 0
        
        ' Iterate through each Stock Transaction

        Do While EndOfData = False
        
            Ticker = Worksheets(DataTab).Cells(DataRow, 1).Value
                  
            If Ticker = PreviousTicker Then
                
                ' Ticker Unchanged
                
                ' Update Volume Total
                TickerSumVolume = TickerSumVolume + Worksheets(DataTab).Cells(DataRow, 7).Value
                
                ' Update the End Price
                EndPrice = Worksheets(DataTab).Cells(DataRow, 6).Value
                
            Else
                                
                ' New Ticker Detected

                If Not PreviousTicker = "" Then
                
                    ' Calculate Statistics Prior Ticker
                    If StartPrice = 0 Then
                        PercentageChange = 0
                    Else
                        PercentageChange = (EndPrice / StartPrice) - 1
                    End If
                        
                    YearlyChange = EndPrice - StartPrice
                    
                    If PercentageChange > TopPerformerValue Then
                        TopPerformerTicker = PreviousTicker
                        TopPerformerValue = PercentageChange
                    End If
                    If PercentageChange < BottomPerformerValue Then
                        BottomPerformerTicker = PreviousTicker
                        BottomPerformerValue = PercentageChange
                    End If
                    If TickerSumVolume > HighestVolumeValue Then
                        HighestVolumeTicker = PreviousTicker
                        HighestVolumeValue = TickerSumVolume
                    End If

                    ' Display the Statistics for the Prior Ticker
                    Worksheets(DataTab).Cells(OutputRow, 12).Value = PreviousTicker
                    Worksheets(DataTab).Cells(OutputRow, 13).Value = YearlyChange
                    Worksheets(DataTab).Cells(OutputRow, 14).Value = PercentageChange
                    Worksheets(DataTab).Cells(OutputRow, 15).Value = TickerSumVolume
                    
                    OutputRow = OutputRow + 1
                    
                End If
                                
                ' Capture data for new TIcker
                TickerSumVolume = Worksheets(DataTab).Cells(DataRow, 7).Value
                StartPrice = Worksheets(DataTab).Cells(DataRow, 3).Value
                EndPrice = Worksheets(DataTab).Cells(DataRow, 6).Value
                
            End If
                     
            PreviousTicker = Ticker
            DataRow = DataRow + 1
            
            If Ticker = "" Then
                EndOfData = True
            End If
        
        Loop
        
        ' Greatest % Increase
        Worksheets(DataTab).Cells(2, 18).Value = TopPerformerTicker
        Worksheets(DataTab).Cells(2, 19).Value = TopPerformerValue
        
        ' Greatest % Decrease
        Worksheets(DataTab).Cells(3, 18).Value = BottomPerformerTicker
        Worksheets(DataTab).Cells(3, 19).Value = BottomPerformerValue
        
        ' Greatest Total Volume
        Worksheets(DataTab).Cells(4, 18).Value = HighestVolumeTicker
        Worksheets(DataTab).Cells(4, 19).Value = HighestVolumeValue
        
   Next DataTab

   MsgBox ("Completed")

End Sub



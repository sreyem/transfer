


#Region "System Imports"

Imports System.ComponentModel
Imports System.IO
Imports System.Xml.Serialization

#End Region


Imports Tools
Imports FOCUSswDB

<Serializable>
<TypeConverter(GetType(PropGridConverter))>
Public Class dryDepositionPv45


    Const StdRegressionCoefficient As Double = -0.05446
    Const StdDigits As Integer = 5

    Const CATCrop As String = "01 Crop"
    Const CATAppln As String = "02 Application"
    Const CATSubstance As String = "03 Substance"
    Const CATDeposition As String = "04 Deposition"
    Const CATSWAN As String = " SWAN"

    Const ShowAll As Boolean = True


#Region "Enums"

    Public Enum eCrops

        Cereals_spring
        Cereals_winter
        Citrus
        Cotton
        Field_beans
        Grass_alfalfa
        Hops
        Legumes
        Maize
        Oil_seed_rape_spring
        Oil_seed_rape_winter
        Olives
        Pome_stone_fruit_early_applns
        Pome_stone_fruit_late_applns
        Potatoes
        Soybeans
        Sugar_beets
        Sunflowers
        Tobacco
        Vegetables_bulb
        Vegetables_fruiting
        Vegetables_leafy
        Vegetables_root
        Vines_early_applns
        Vines_late_applns
        notDefined

    End Enum

    Public Enum eEVACropGroup

        orcharding_early
        orcharding_late
        vine
        agriculture
        agriculture_2
        agriculture_larger_than_900_lperha
        allotment_patch_smaller_than_50cm
        allotment_patch_taler_than_50cm
        allotment_tree_late
        allotment_tree_early_smaller_than_2m
        allotment_tree_early_taler_than_2m
        knapsack_sprayer
        hop_growing


    End Enum

    Public Enum eAreaOrVolumeCulture
        Area
        Volume
    End Enum

    Public Enum eEVAClassification

        non_volatile
        semi_low
        semi_medium
        volatile

    End Enum

#End Region

#Region "DBs"

    <XmlIgnore>
    <Browsable(False)>
    Public Property VPClassificationDB As String() =
      {"non volatile =           VP <  1*10-5",
       "semi low     = 1*10-5 <= VP >  1*10-4",
       "semi medium  = 1*10-4 <  VP >  5*10-3",
       "volatile     =           VP >= 5*10-3"}

    <XmlIgnore>
    <Browsable(False)>
    Public PercDepositionAt1mDB As Double() =
        {
            0,
            0.089,
            0.221,
            1.555
        }

    <XmlIgnore>
    <Browsable(False)>
    Private PercApplnRateRelevantForDepositionDB As Double() =
         {
             97.56569,
             98.98412,
             99.58267,
             99.74808,
             99.49616,
             99.8068242,
             99.98161,
             99.987595,
             99.95265,
             99.82215,
             99.383,
             99.99709,
             98.62724
         }

    <XmlIgnore>
    <Browsable(False)>
    Private DepositionRatePercDB As Double(,) =
       {
           {9.091, 16.67},
           {9.091, 8.333},
           {9.091, 8.333},
           {9.091, 8.333},
           {4.545, 4.167},
           {4.545, 4.167},
           {4.545, 4.167},
           {4.545, 4.167},
           {4.545, 4.167},
           {4.545, 4.167},
           {4.545, 4.167},
           {4.545, 4.167},
           {2.273, 2.083},
           {2.273, 2.083},
           {2.273, 2.083},
           {2.273, 2.083},
           {2.273, 2.083},
           {2.273, 2.083},
           {2.273, 2.083},
           {2.273, 2.083},
           {2.273, 2.083},
           {2.273, 2.083},
           {2.273, 2.083},
           {2.273, 2.083}
       }


#End Region

#Region "01 Crop"

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private m_EVACropGroup As eEVACropGroup = eEVACropGroup.agriculture

    ''' <summary>
    ''' EVA Crop Group
    ''' </summary>
    <Category(CATCrop)>
    <DisplayName("EVA Crop Group")>
    <Description("Volume: orcharding (early/late), vine, hop, allotment tree(early/late)" & vbCrLf &
                 "Area  : agriculture /*2/>900l/ha ,allotment patch <> 50cm,knapsack sprayer")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(ShowAll)>
    <[ReadOnly](False)>
    <DefaultValue(eEVACropGroup.agriculture)>
    Public Property EVACropGroup As eEVACropGroup
        Get
            Return m_EVACropGroup
        End Get
        Set(vEVACropGroup As eEVACropGroup)
            m_EVACropGroup = vEVACropGroup
        End Set
    End Property



    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private m_AreaOrVolumeCulture As eAreaOrVolumeCulture = eAreaOrVolumeCulture.Area

    ''' <summary>
    ''' Area Or Volume Culture
    ''' Volume: orcharding (early/late), vine, hop, allotment tree(early/late)
    ''' Area  : agriculture /*2/>900l/ha ,allotment patch greater/smaller 50cm,knapsack sprayer
    ''' </summary>
    <Category(CATCrop)>
    <DisplayName("Area Or Volume Culture")>
    <Description("Volume: orcharding (early/late), vine, hop, allotment tree(early/late)" & vbCrLf &
                 "Area  : agriculture /*2/>900l/ha ,allotment patch <> 50cm,knapsack sprayer")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(ShowAll)>
    <[ReadOnly](False)>
    Public ReadOnly Property AreaOrVolumeCulture As eAreaOrVolumeCulture
        Get
            Return getAreaOrVolumeCulture(Me.EVACropGroup)
        End Get
    End Property


    Private Function getAreaOrVolumeCulture(EVACropGroup As eEVACropGroup) As eAreaOrVolumeCulture

        Select Case EVACropGroup

            Case eEVACropGroup.agriculture,
                 eEVACropGroup.agriculture_2,
                 eEVACropGroup.agriculture_larger_than_900_lperha,
                 eEVACropGroup.knapsack_sprayer

                Return eAreaOrVolumeCulture.Area

            Case Else

                Return eAreaOrVolumeCulture.Volume

        End Select

    End Function

#End Region

#Region "02 Application"

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private m_ApplnRate As Double = Double.NaN

    ''' <summary>
    ''' Application rates in g/ha
    ''' </summary>
    <Category(CATAppln)>
    <DisplayName("Appln Rate")>
    <Description("Application rate in g/ha" & vbCrLf &
                 "")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(ShowAll)>
    <[ReadOnly](False)>
    Public Property ApplnRate As Double
        Get
            Return m_ApplnRate
        End Get
        Set(vApplnRate As Double)
            m_ApplnRate = vApplnRate
        End Set
    End Property


    Private Function getDepositionRatePerc(
                                    Hour As Integer,
                                    AreaOrVolumeCulture As eAreaOrVolumeCulture) As Double
        Return DepositionRatePercDB(Hour, AreaOrVolumeCulture)
    End Function



    ''' <summary>
    ''' Application rate corrected for deposition
    ''' (appln. rate - drift)
    ''' </summary>
    <Category(CATAppln)>
    <DisplayName("Appln Rate Relevant For Deposition")>
    <Description("Application rate corrected for deposition" & vbCrLf &
                 "(appln. rate - drift)")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(ShowAll)>
    <[ReadOnly](False)>
    Public ReadOnly Property ApplnRateRelevantForDeposition As Double
        Get
            Return getApplnRateRelevantForDeposition(Me.ApplnRate)
        End Get
    End Property

    Private Function getApplnRateRelevantForDeposition(
                                                 ApplnRate As Double,
                                        Optional Digits As Integer = StdDigits) As Double

        Return Math.Round(ApplnRate * PercApplnRateRelevantForDepositionDB(EVACropGroup) /
                                        100,
                     digits:=Digits)

    End Function


#End Region

#Region "03 Substance"

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private m_EVAClassification As eEVAClassification = eEVAClassification.non_volatile

    ''' <summary>
    ''' EVAClassification
    ''' </summary>
    <Category(CATSubstance)>
    <DisplayName("Classification")>
    <Description("Classification due to VP" & vbCrLf &
                 "non volatile, semi volatile low/medium(std) or volatile")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(ShowAll)>
    <[ReadOnly](False)>
    Public Property EVAClassification As eEVAClassification
        Get
            Return m_EVAClassification
        End Get
        Set(vEVAClassification As eEVAClassification)
            m_EVAClassification = vEVAClassification
        End Set
    End Property


    ''' <summary>
    ''' VPGroups
    ''' </summary>
    <Category(CATSubstance)>
    <DisplayName("Vapor Pressure Class")>
    <Description("Vapor Pressure in Pa " & vbCrLf &
                 "")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(ShowAll)>
    <[ReadOnly](False)>
    Public ReadOnly Property VPGroups As String
        Get
            Return VPClassificationDB(FOCUSEVAClassification)
        End Get
    End Property


    ''' <summary>
    ''' Deposition value in 1m distance in percent
    ''' due to classification
    ''' </summary>
    <Category(CATSubstance)>
    <DisplayName("Deposition in 1m")>
    <Description("Deposition value in 1m distance in percent" & vbCrLf &
                 "due to classification")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(ShowAll)>
    <[ReadOnly](False)>
    Public ReadOnly Property PercDeposition1m As Double
        Get
            Return PercDepositionAt1mDB(FOCUSEVAClassification)
        End Get
    End Property

#End Region

#Region "04 Deposition"

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private m_BufferDistance As Integer = 1

    ''' <summary>
    ''' Downwind distance from the treated crop in m
    ''' </summary>
    <Category(CATDeposition)>
    <DisplayName("Buffer Distance")>
    <Description("Downwind distance from the treated crop in m" & vbCrLf &
                 "")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(ShowAll)>
    <[ReadOnly](False)>
    Public Property BufferDistance As Integer
        Get
            Return m_BufferDistance
        End Get
        Set(vBufferDistance As Integer)
            m_BufferDistance = vBufferDistance
        End Set
    End Property

    ''' <summary>
    ''' Deposition Buffer
    ''' </summary>
    <Category(CATDeposition)>
    <DisplayName("Deposition in Buffer Distance")>
    <Description("in percent" & vbCrLf &
                 "")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(ShowAll)>
    <[ReadOnly](False)>
    Public ReadOnly Property PercDepositionInXm As Double
        Get

            Return calcPercDepositionInXm(PercDepositionIn1m:=Me.PercDeposition1m,
                                     AreaOrVolumeCulture:=Me.AreaOrVolumeCulture,
                                          BufferDistance:=Me.BufferDistance)

        End Get
    End Property

    Private Function calcPercDepositionInXm(
                                         PercDepositionIn1m As Double,
                                         BufferDistance As Double,
                                         AreaOrVolumeCulture As eAreaOrVolumeCulture,
                                Optional RegressionCoefficient As Double = StdRegressionCoefficient,
                                Optional Digits As Integer = StdDigits) As Double


        Return Math.Round((AreaOrVolumeCulture + 1) *
                          PercDepositionIn1m *
                          Math.Exp(RegressionCoefficient * (BufferDistance - 1)),
                    digits:=5)

    End Function



    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private m_PlantInterception As Double = Double.NaN

    ''' <summary>
    ''' Plant interception in percent 0 -100
    ''' </summary>
    <Category(CATDeposition)>
    <DisplayName("Plant Interception")>
    <Description("Plant interception in percent" & vbCrLf &
                 "0 -100")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(ShowAll)>
    <[ReadOnly](False)>
    Public Property PlantInterception As Double
        Get
            Return m_PlantInterception
        End Get
        Set(vPlantInterception As Double)
            m_PlantInterception = vPlantInterception
        End Set
    End Property

#Region "Deposition from Plant"


    ''' <summary>
    ''' Deposition after volatilization from plant surfaces
    ''' </summary>
    <Category(CATDeposition)>
    <DisplayName("Deposition from Plant")>
    <Description("Deposition after volatilization" & vbCrLf &
                 "from plant surfaces")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(ShowAll)>
    <[ReadOnly](False)>
    Public ReadOnly Property DepositionFromPlant As Double
        Get
            Return calcDepositionFromPlant(ApplnRateRelevantForDeposition:=Me.ApplnRateRelevantForDeposition,
                                       PercDepositionXm:=Me.PercDepositionInXm,
                                       PlantInterception:=Me.PlantInterception)
        End Get
    End Property

    Private Function calcDepositionFromPlant(
                                ApplnRateRelevantForDeposition As Double,
                                PlantInterception As Double,
                                PercDepositionXm As Double,
                       Optional Digits As Integer = 5) As Double

        Return Math.Round(ApplnRateRelevantForDeposition * PercDepositionXm * PlantInterception /
                                        10000,
                    digits:=Digits)

    End Function


#End Region

#Region "Deposition from Soil"


    ''' <summary>
    ''' Deposition after volatilization from soil surfaces
    ''' </summary>
    <Category(CATDeposition)>
    <DisplayName("Deposition from Soil")>
    <Description("Deposition after volatilization" & vbCrLf &
                 "from soil surfaces")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(ShowAll)>
    <[ReadOnly](False)>
    Public ReadOnly Property DepositionFromSoil As Double
        Get
            Return calcDepositionFromSoil(
                        ApplnRateRelevantForDeposition:=Me.ApplnRateRelevantForDeposition,
                        PercDepositionXm:=Me.PercDepositionInXm,
                        PlantInterception:=Me.PlantInterception)
        End Get
    End Property


    Private Function calcDepositionFromSoil(
                                         ApplnRateRelevantForDeposition As Double,
                                         PlantInterception As Double,
                                         PercDepositionXm As Double,
                                Optional Digits As Integer = 5) As Double

        Return Math.Round(ApplnRateRelevantForDeposition * PercDepositionXm * (100 - PlantInterception) /
                                             30000,
                    digits:=Digits)

    End Function




#End Region



    ''' <summary>
    ''' Deposition Rates
    ''' </summary>
    <Category(CATDeposition)>
    <DisplayName("Deposition Rates")>
    <Description("Deposition rates after different times" & vbCrLf &
                 "in mg/m2")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(ShowAll)>
    <[ReadOnly](False)>
    Public ReadOnly Property DepositionRates As Double()
        Get
            Return calcDepositionRates(AreaOrVolumeCulture:=Me.AreaOrVolumeCulture,
                                       DepositionFromPlant:=Me.DepositionFromPlant,
                                       DepositionFromSoil:=Me.DepositionFromSoil)
        End Get
    End Property


    Private Function calcDepositionRates(
                                AreaOrVolumeCulture As eAreaOrVolumeCulture,
                                DepositionFromPlant As Double,
                                DepositionFromSoil As Double) As Double()

        Dim Result As New List(Of Double)
        Dim DepositionRatePerc As Double = Double.NaN


        For Hour As Integer = 0 To 23

            DepositionRatePerc = DepositionRatePercDB(Hour, AreaOrVolumeCulture)

            Result.Add(Math.Round(DepositionRatePerc *
                                  (DepositionFromPlant + DepositionFromSoil) / 1000, digits:=StdDigits))

            If Result.Last = Double.NaN Then
                Result(Result.Count - 1) = 0
            End If


        Next


        Return Result.ToArray

    End Function


#End Region

#Region "05 SWAN"


    <Category(CATSWAN)>
    <RefreshProperties(RefreshProperties.All)>
    Public Property PMTDescription As String() = {}

    Private m_FOCUSBuffer As Double = 1

    <Category(CATSWAN)>
    <RefreshProperties(RefreshProperties.All)>
    Public Property FOCUSBuffer As Double
        Get
            Return m_FOCUSBuffer
        End Get
        Set(value As Double)
            m_FOCUSBuffer = value
            RaiseEvent recreateSWANDryDepositionMatrix()
        End Set
    End Property

    Private m_FOCUSApplnRates As Double() = {}

    <Category(CATSWAN)>
    <RefreshProperties(RefreshProperties.All)>
    Public Property FOCUSApplnRates As Double()
        Get
            Return m_FOCUSApplnRates
        End Get
        Set(value As Double())
            m_FOCUSApplnRates = value
            RaiseEvent recreateSWANDryDepositionMatrix()
        End Set
    End Property

    Private m_FOCUSStep34Crop As eCrops = eCrops.Cereals_winter

    <Category(CATSWAN)>
    <RefreshProperties(RefreshProperties.All)>
    Public Property FOCUSStep34Crop As eCrops
        Get
            Return m_FOCUSStep34Crop
        End Get
        Set(value As eCrops)
            m_FOCUSStep34Crop = value
            RaiseEvent recreateSWANDryDepositionMatrix()
        End Set
    End Property

    Private m_FOCUSPlantInterception As Double = Double.NaN

    <Category(CATSWAN)>
    <RefreshProperties(RefreshProperties.All)>
    Public Property FOCUSPlantInterception As Double
        Get
            Return m_FOCUSPlantInterception
        End Get
        Set(value As Double)
            m_FOCUSPlantInterception = value
            RaiseEvent recreateSWANDryDepositionMatrix()
        End Set
    End Property

    Private m_FOCUSEVAClassification As eEVAClassification = eEVAClassification.non_volatile

    <Category(CATSWAN)>
    <RefreshProperties(RefreshProperties.All)>
    Public Property FOCUSEVAClassification As eEVAClassification
        Get
            Return m_FOCUSEVAClassification
        End Get
        Set(value As eEVAClassification)
            m_FOCUSEVAClassification = value
            RaiseEvent recreateSWANDryDepositionMatrix()
        End Set
    End Property

    <Category(CATSWAN)>
    <RefreshProperties(RefreshProperties.All)>
    Public Property CalcEVADryDepositionMatrix As Boolean
        Get
            Return True
        End Get
        Set(value As Boolean)
            If value = True Then Exit Property

            SWANDryDepositionMatrix = createSWANDryDepositionMatrix(
                   FOCUSApplnRates:=Me.FOCUSApplnRates,
                       FOCUSBuffer:=Me.FOCUSBuffer,
                   FOCUSStep34Crop:=Me.FOCUSStep34Crop,
            FOCUSPlantInterception:=Me.FOCUSPlantInterception,
                 EVAClassification:=Me.FOCUSEVAClassification)

        End Set
    End Property

    <Category(CATSWAN)>
    <RefreshProperties(RefreshProperties.All)>
    Public Property SWANDryDepositionMatrix As String() =
        {
            "0-1    0 0 0 0 0 0 0 0",
            "1-2    0 0 0 0 0 0 0 0",
            "2-3    0 0 0 0 0 0 0 0",
            "3-4    0 0 0 0 0 0 0 0",
            "4-5    0 0 0 0 0 0 0 0",
            "5-6    0 0 0 0 0 0 0 0",
            "6-7    0 0 0 0 0 0 0 0",
            "7-8    0 0 0 0 0 0 0 0",
            "8-9    0 0 0 0 0 0 0 0",
            "9-10   0 0 0 0 0 0 0 0",
            "10-11  0 0 0 0 0 0 0 0",
            "11-12  0 0 0 0 0 0 0 0",
            "12-13  0 0 0 0 0 0 0 0",
            "13-14  0 0 0 0 0 0 0 0",
            "14-15  0 0 0 0 0 0 0 0",
            "15-16  0 0 0 0 0 0 0 0",
            "16-17  0 0 0 0 0 0 0 0",
            "17-18  0 0 0 0 0 0 0 0",
            "18-19  0 0 0 0 0 0 0 0",
            "19-20  0 0 0 0 0 0 0 0",
            "20-21  0 0 0 0 0 0 0 0",
            "21-22  0 0 0 0 0 0 0 0",
            "22-23  0 0 0 0 0 0 0 0",
            "23-24  0 0 0 0 0 0 0 0"
        }



    Private Function createSWANDryDepositionMatrix(
                                                   FOCUSApplnRates As Double(),
                                                   FOCUSBuffer As Double,
                                                   FOCUSStep34Crop As eCrops,
                                                   FOCUSPlantInterception As Double,
                                                   EVAClassification As eEVAClassification) As String()

        Dim Result As New List(Of String)
        Dim temprow As String = ""
        Dim AreaOrVolumeCulture As eAreaOrVolumeCulture
        Dim DepositionRates As New List(Of Double())
        Dim temp As Double() = {}
        Dim DepositionFromPlant As Double = Double.NaN
        Dim DepositionFromSoil As Double = Double.NaN

        Dim ApplnRateRelevantForDeposition As Double = Double.NaN
        Dim PercDepositionXm As Double = Double.NaN

        Const Delimiter As String = " "

        Select Case FOCUSStep34Crop

            Case eCrops.Citrus,
                 eCrops.Hops,
                 eCrops.Olives,
                 eCrops.Pome_stone_fruit_early_applns,
                 eCrops.Pome_stone_fruit_late_applns,
                 eCrops.Vines_early_applns,
                 eCrops.Vines_late_applns

                AreaOrVolumeCulture = eAreaOrVolumeCulture.Volume

            Case Else

                AreaOrVolumeCulture = eAreaOrVolumeCulture.Area

        End Select



        For Each FOCUSApplnRate As Double In FOCUSApplnRates


            ApplnRateRelevantForDeposition = getApplnRateRelevantForDeposition(
                                                ApplnRate:=FOCUSApplnRate)

            PercDepositionXm = calcPercDepositionInXm(
                                     AreaOrVolumeCulture:=AreaOrVolumeCulture,
                                     BufferDistance:=FOCUSBuffer,
                                     PercDepositionIn1m:=PercDepositionAt1mDB(EVAClassification),
                                     RegressionCoefficient:=StdRegressionCoefficient)

            DepositionFromPlant = calcDepositionFromPlant(
                                       ApplnRateRelevantForDeposition:=ApplnRateRelevantForDeposition,
                                       PlantInterception:=FOCUSPlantInterception,
                                       PercDepositionXm:=PercDepositionXm)

            DepositionFromSoil = calcDepositionFromSoil(
                                       ApplnRateRelevantForDeposition:=ApplnRateRelevantForDeposition,
                                       PlantInterception:=FOCUSPlantInterception,
                                       PercDepositionXm:=PercDepositionXm)


            temp = calcDepositionRates(
                            AreaOrVolumeCulture:=AreaOrVolumeCulture,
                            DepositionFromPlant:=DepositionFromPlant,
                             DepositionFromSoil:=DepositionFromSoil)


            DepositionRates.Add(temp)

        Next

        For HourCounter As Integer = 0 To 23

            temprow = (HourCounter & "-" & HourCounter + 1).PadRight("0-1    ".Length)

            For ApplnCounter As Integer = 0 To 7

                If FOCUSApplnRates.Count - 1 >= ApplnCounter Then

                    If FOCUSEVAClassification = eEVAClassification.non_volatile OrElse
                       Double.IsNaN(DepositionRates(ApplnCounter)(HourCounter)) Then

                        temprow &= "0" & Delimiter
                    Else
                        temprow &= DepositionRates(ApplnCounter)(HourCounter).ToString("0.00000") & Delimiter

                    End If



                Else

                    temprow &= "0"

                    If ApplnCounter <> 7 Then
                        temprow &= Delimiter
                    End If

                End If

            Next

            Result.Add(temprow)


        Next

        Return Result.ToArray

    End Function


    Private Event recreateSWANDryDepositionMatrix()

    Private Sub DryDeposition_recreateSWANDryDepositionMatrix() Handles Me.recreateSWANDryDepositionMatrix
        Me.CalcEVADryDepositionMatrix = False
    End Sub


#End Region

End Class


Public Class dryDeposition

    Public Sub New()

    End Sub

    Public Const catCrop As String = "01  Crop"

#Region "    Crop"

    Private _FOCUSswDriftCrop As eFOCUSswDriftCrop = eFOCUSswDriftCrop.not_defined
    Private _cultureType As eCultureType = eCultureType.not_defined
    Private _vaporPressure As Double = stdDblEmpty
    Private _evaClassification As eEVAClassification = eEVAClassification.not_defined
    Private _percDeposition1m As Double = stdDblEmpty
    Private _relevant4Deposition As Double = stdDblEmpty
    Private _EVAcropGroup As eEVACropGroup = eEVACropGroup.not_defined

    ''' <summary>
    ''' FOCUSsw Crop as enum
    ''' </summary>
    <Category(catCrop)>
    <RefreshProperties(RefreshProperties.All)>
    <DisplayName(
        "Crop")>
    <Description(
        "FOCUSsw Crop 2 digit PRZM" & vbCrLf &
        "crop code, e.g. MZ = Maize")>
    <Browsable(True)>
    <[ReadOnly](False)>
    <DefaultValue("")>
    Public Property FOCUSswDriftCrop As eFOCUSswDriftCrop
        Get
            Return _FOCUSswDriftCrop
        End Get
        Set
            _FOCUSswDriftCrop = Value
            Me.cultureType = getCultureType(FOCUSswDriftCrop:=_FOCUSswDriftCrop.ToString)
        End Set
    End Property

    ''' <summary>
    ''' Enum Culture type
    ''' area (Maize, Cereals etc.) or volume (Vines, Olives etc.)
    ''' </summary>
    <TypeConverter(GetType(EnumConverter(Of eCultureType)))>
    Public Enum eCultureType

        ''' <summary>
        ''' start setting : not def.
        ''' </summary>
        <Description(EnumConverter(Of Type).not_defined)>
        not_defined = -1

        ''' <summary>
        ''' Area (Maize, Cereals etc.)
        ''' </summary>
        <Description("Area")>
        Area

        ''' <summary>
        ''' Area culture
        ''' </summary>
        <Description("Volume")>
        Volume

    End Enum

    ''' <summary>
    ''' Get culture type (area or volume) from FOCUSsw drift crop
    ''' </summary>
    ''' <param name="FOCUSswDriftCrop">
    ''' FOCUSsw Crop 2 digit PRZM
    ''' crop code, e.g. MZ = Maize
    ''' </param>
    ''' <returns>
    ''' culture type (area or volume) as eCultureType
    ''' </returns>
    Private Function getCultureType(FOCUSswDriftCrop As String) As eCultureType

        Select Case FOCUSswDriftCrop

            Case eFOCUSswDriftCrop.not_defined.ToString

                Return eCultureType.not_defined

            Case eFOCUSswDriftCrop.PE.ToString, eFOCUSswDriftCrop.PL.ToString,
                 eFOCUSswDriftCrop.HP.ToString,
                 eFOCUSswDriftCrop.CI.ToString,
                 eFOCUSswDriftCrop.OL.ToString,
                 eFOCUSswDriftCrop.VI.ToString, eFOCUSswDriftCrop.VA.ToString

                Return eCultureType.Volume

            Case Else

                Return eCultureType.Area

        End Select


    End Function

    ''' <summary>
    ''' Culture type
    ''' area (Maize, Cereals etc.) or volume (Vines, Olives etc.)
    ''' </summary>
    <Category(catCrop)>
    <RefreshProperties(RefreshProperties.All)>
    <DisplayName(
"Culture Type")>
    <Description(
    "Culture type e.g. area (Maize, Cereals etc.)" & vbCrLf &
    "or volume (Vines, Olives etc.)")>
    <Browsable(True)>
    <[ReadOnly](False)>
    <DefaultValue(CInt(eCultureType.not_defined))>
    Public Property cultureType As eCultureType
        Get
            Return _cultureType
        End Get
        Set
            _cultureType = Value
            vdPlants = calcVDPlants(CultureType:=_cultureType, Buffer:=1)
            vdSoil = calcVDSoil(CultureType:=_cultureType, EVAClassification:=Me.evaClassification)
        End Set
    End Property

    <TypeConverter(GetType(EnumConverter(Of eEVACropGroup)))>
    Public Enum eEVACropGroup

        ''' <summary>
        ''' start setting : not def.
        ''' </summary>
        <Description(EnumConverter(Of Type).not_defined)>
        not_defined = -1

        <Description("arable crops")>
        arable_crops

        <Description("arable crops × 2")>
        agriculture_2

        <Description("arable crops >900 L/ha")>
        agriculture_larger_than_900_lperha

        <Description("vines")>
        vine

        <Description("orchards early")>
        orcharding_early

        <Description("orchards late")>
        orcharding_late

        <Description("hops")>
        hops

        <Description("knapsack sprayer, plants <50 cm")>
        allotment_patch_smaller_than_50cm

        <Description("knapsack sprayer, plants >50 cm")>
        allotment_patch_taler_than_50cm

        <Description("knapsack sprayer, spraying screen")>
        allotment_tree_late

        <Description("gardening, plants <50 cm")>
        gardening_plants_smaller_than_50cm

        <Description("gardening, plants >50 cm")>
        gardening_plants_taler_than_50cm

        <Description("gardening, trees early <2 m")>
        allotment_tree_early_smaller_than_2m

        <Description("gardening, trees early >2 m")>
        allotment_tree_early_taler_than_2m

        <Description("gardening, trees late")>
        gardening_trees_late

        <Description("greenhouse (no drift)")>
        greenhouse_no_drift


    End Enum


    Public Function calcVDPlants(CultureType As eCultureType, ByRef Buffer As Integer) As Double



        Dim cropHeightFactor As Integer = 1
        Const expFactor As Double = -0.05446

        If CultureType = eCultureType.Volume Then

            cropHeightFactor = 2

            If Buffer < 3 Then

                Buffer = 3

            End If
        End If

        Return percDeposition1m * Math.Exp(expFactor * (Buffer - 1)) * cropHeightFactor

    End Function

    <DefaultValue(stdDblEmpty)>
    <TypeConverter(GetType(DblConv))>
    <AttributeProvider(
    "unit='%'")>
    Public Property vdPlants As Double = stdDblEmpty


    Public Function calcVDSoil(CultureType As eCultureType, EVAClassification As eEVAClassification) As Double

        If vdPlants = stdDblEmpty Then Return stdDblEmpty
        If EVAClassification = eEVAClassification.semi_plants Then Return 0

        Dim cropHeightFactor As Integer = 1

            If CultureType = eCultureType.Volume Then
                cropHeightFactor = 2
            End If

            Return vdPlants / 3 / cropHeightFactor

    End Function

    <DefaultValue(stdDblEmpty)>
    <TypeConverter(GetType(DblConv))>
    <AttributeProvider(
    "unit='%'")>
    Public Property vdSoil As Double = stdDblEmpty

#End Region

    Public Const catEVA As String = "02  EVA"

#Region "    EVA Class"

    <TypeConverter(GetType(EnumConverter(Of eEVAClassification)))>
    Public Enum eEVAClassification

        ''' <summary>
        ''' start setting : not def.
        ''' </summary>
        <Description(EnumConverter(Of Type).not_defined)>
        not_defined = -1

        <Description("non volatile      = " & vbCrLf & "          VP <  1*10-5")>
        non_volatile

        <Description("semi plants       = " & vbCrLf & "1*10-5 <= VP >  1*10-4")>
        semi_plants

        <Description("semi plants/soil  = " & vbCrLf & "1*10-4 <  VP >  5*10-3")>
        semi_plants_soil

        <Description("non volatile      = " & vbCrLf & "          VP >= 5*10-3")>
        volatile

    End Enum

    ''' <summary>
    ''' Vapor Pressure in PA
    ''' to specify EVA classification
    ''' </summary>
    ''' <returns></returns>
    <Category(catEVA)>
    <RefreshProperties(RefreshProperties.All)>
    <DisplayName(
"Vapor Pressure")>
    <Description(
    "Vapor Pressure in PA" & vbCrLf &
    "to specify EVA classification")>
    <Browsable(True)>
    <[ReadOnly](False)>
    <DefaultValue(stdDblEmpty)>
    <TypeConverter(GetType(DblConv))>
    <AttributeProvider(
    "format='0.0E-00'")>
    Public Property vaporPressure As Double
        Get
            Return _vaporPressure
        End Get
        Set
            _vaporPressure = Value
            Me.evaClassification = getEVAClass(vaporPressure:=_vaporPressure)
        End Set
    End Property

    Private Function getEVAClass(vaporPressure As Double) As eEVAClassification

        Const nonVolatile As Double = 0.00001
        Const semiLow As Double = 0.0001
        Const semiMedium As Double = 0.005

        If vaporPressure < nonVolatile Then

            Return eEVAClassification.non_volatile

        ElseIf vaporPressure >= nonVolatile AndAlso
               vaporPressure < semiLow Then

            Return eEVAClassification.semi_plants

        ElseIf vaporPressure >= semiLow AndAlso
               vaporPressure < semiMedium Then

            Return eEVAClassification.semi_plants_soil

        Else

            Return eEVAClassification.volatile

        End If

    End Function

    ''' <summary>
    ''' Classification according to EVA
    ''' based on vapor pressure: non volatile, low, medium and volatile
    ''' </summary>
    ''' <returns></returns>
    <Category(catEVA)>
    <RefreshProperties(RefreshProperties.All)>
    <DisplayName(
"EVA Class")>
    <Description(
    "Classification according to EVA " & vbCrLf &
    "based on vapor pressure: non volatile, low, medium and volatile")>
    <Browsable(True)>
    <[ReadOnly](False)>
    <DefaultValue(CInt(eCultureType.not_defined))>
    Public Property evaClassification As eEVAClassification
        Get
            Return _evaClassification
        End Get
        Set
            _evaClassification = Value
            percDeposition1m = percDepositionAt1mDB(_evaClassification)
            relevant4Deposition = percApplnRateRelevantForDepositionDB(_evaClassification)
        End Set
    End Property

    ''' <summary>
    ''' Deposition value in 1m distance in percent
    ''' due to classification
    ''' </summary>
    <Category(catEVA)>
    <DisplayName(
    "Deposition in 1m")>
    <Description(
    "Deposition value in 1m distance in percent" & vbCrLf &
    "due to classification")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(True)>
    <[ReadOnly](False)>
    <DefaultValue(stdDblEmpty)>
    <TypeConverter(GetType(DblConv))>
    <AttributeProvider(
    "unit='%'")>
    Public Property percDeposition1m As Double
        Get
            Return _percDeposition1m
        End Get
        Set
            _percDeposition1m = Value
            vdPlants = calcVDPlants(CultureType:=_cultureType, Buffer:=1)
            vdSoil = calcVDSoil(CultureType:=_cultureType, EVAClassification:=Me.evaClassification)
        End Set
    End Property

    <XmlIgnore>
    <Browsable(False)>
    Public percDepositionAt1mDB As Double() =
        {
            0,
            0.089,
            0.221,
            1.555
        }

    <Category(catEVA)>
    <DisplayName(
    "Relevant for Deposition")>
    <Description(
    "Percent of appln rate " & vbCrLf &
    "relevant for deposition")>
    <RefreshProperties(RefreshProperties.All)>
    <DebuggerBrowsable(DebuggerBrowsableState.Collapsed)>
    <Browsable(True)>
    <[ReadOnly](False)>
    <DefaultValue(stdDblEmpty)>
    <TypeConverter(GetType(DblConv))>
    <AttributeProvider(
    "unit='%'")>
    Public Property relevant4Deposition As Double
        Get
            Return _relevant4Deposition
        End Get
        Set
            _relevant4Deposition = Value
        End Set
    End Property

    <XmlIgnore>
    <Browsable(False)>
    Private percApplnRateRelevantForDepositionDB As Double() =
         {
             97.56569,
             98.98412,
             99.58267,
             99.74808,
             99.49616,
             99.8068242,
             99.98161,
             99.987595,
             99.95265,
             99.82215,
             99.383,
             99.99709,
             98.62724
         }


#End Region

#Region "    Appln Rate"


    Public Property dryDepositionDay As dryDepositionDay() = {}


#End Region

End Class

<TypeConverter(GetType(PropGridConverter))>
<Serializable>
Public Class dryDepositionDay

    Public Sub New()

    End Sub

    Public Property rate As Double

    Public Property interception As Double






    <XmlIgnore>
    <Browsable(False)>
    Private PercApplnRateRelevantForDepositionDB As Double() =
         {
             97.56569,
             98.98412,
             99.58267,
             99.74808,
             99.49616,
             99.8068242,
             99.98161,
             99.987595,
             99.95265,
             99.82215,
             99.383,
             99.99709,
             98.62724
         }

    <XmlIgnore>
    <Browsable(False)>
    Private DepositionRatePercDB As Double(,) =
       {
           {9.091, 16.67},
           {9.091, 8.333},
           {9.091, 8.333},
           {9.091, 8.333},
           {4.545, 4.167},
           {4.545, 4.167},
           {4.545, 4.167},
           {4.545, 4.167},
           {4.545, 4.167},
           {4.545, 4.167},
           {4.545, 4.167},
           {4.545, 4.167},
           {2.273, 2.083},
           {2.273, 2.083},
           {2.273, 2.083},
           {2.273, 2.083},
           {2.273, 2.083},
           {2.273, 2.083},
           {2.273, 2.083},
           {2.273, 2.083},
           {2.273, 2.083},
           {2.273, 2.083},
           {2.273, 2.083},
           {2.273, 2.083}
       }


    Public Property test As Double() =
        {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}


End Class

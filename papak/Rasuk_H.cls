VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Rasuk_H"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private ONE_AtoE, TWO_AtoE, THREE_BtoE

Private FOUR_CtoE, FIVE_DtoE
'''''''''''
Private A_ONEtoTWO, B_ONEtoTHREE, C_ONEtoFOUR

Private D_ONEtoFIVE, E_ONEtoFIVE
'''''''''''

'''''''''''

Public Function H_ONE_AtoE()
   H_ONE_AtoE = ONE_AtoE
End Function
Public Property Let Value20(Setting)
   ONE_AtoE = Setting
End Property

Public Function H_TWO_AtoE()
   H_TWO_AtoE = TWO_AtoE
End Function
Public Property Let Value21(Setting)
   TWO_AtoE = Setting
End Property

Public Function H_THREE_BtoE()
   H_THREE_BtoE = THREE_BtoE
End Function
Public Property Let Value22(Setting)
   THREE_BtoE = Setting
End Property

Public Function H_FOUR_CtoE()
   H_FOUR_CtoE = FOUR_CtoE
End Function
Public Property Let Value23(Setting)
   FOUR_CtoE = Setting
End Property

Public Function H_FIVE_DtoE()
   H_FIVE_DtoE = FIVE_DtoE
End Function
Public Property Let Value24(Setting)
   FIVE_DtoE = Setting
End Property
''''''''''''''''''''''''''''''''''''



Public Function H_A_ONEtoTWO()
   H_A_ONEtoTWO = A_ONEtoTWO
End Function
Public Property Let Value25(Setting)
  A_ONEtoTWO = Setting
End Property

Public Function H_B_ONEtoTHREE()
   H_B_ONEtoTHREE = B_ONEtoTHREE
End Function
Public Property Let Value26(Setting)
  B_ONEtoTHREE = Setting
End Property

Public Function H_C_ONEtoFOUR()
   H_C_ONEtoFOUR = C_ONEtoFOUR
End Function
Public Property Let Value27(Setting)
   C_ONEtoFOUR = Setting
End Property

Public Function H_D_ONEtoFIVE()
   H_D_ONEtoFIVE = D_ONEtoFIVE
End Function
Public Property Let Value28(Setting)
   D_ONEtoFIVE = Setting
End Property

Public Function H_E_ONEtoFIVE()
   H_E_ONEtoFIVE = E_ONEtoFIVE
End Function
Public Property Let Value29(Setting)
   E_ONEtoFIVE = Setting
End Property


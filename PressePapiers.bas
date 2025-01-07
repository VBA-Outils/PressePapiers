Attribute VB_Name = "PressePapiers"
'
' Copier/Coller dans le presse-papiers de Windows - VBA
' https://github.com/VBA-Outils/PressePapiers
'
' @Module PressePapiers
' @author vincent.rosset@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'
' Copyright (c) 2024, Vincent ROSSET
' All rights reserved.
'
' Redistribution and use in source and binary forms, with or without
' modification, are permitted provided that the following conditions are met:
'     * Redistributions of source code must retain the above copyright
'       notice, this list of conditions and the following disclaimer.
'     * Redistributions in binary form must reproduce the above copyright
'       notice, this list of conditions and the following disclaimer in the
'       documentation and/or other materials provided with the distribution.
'     * Neither the name of the <organization> nor the
'       names of its contributors may be used to endorse or promote products
'       derived from this software without specific prior written permission.
'
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
' ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
' WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
' DISCLAIMED. IN NO EVENT SHALL <COPYRIGHT HOLDER> BE LIABLE FOR ANY
' DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
' (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
' LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
' ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
' (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
' SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

' *----------------------------------------------------------------------------------------------------------*
' * Copier/coller une chaîne de caractères depuis/dans le presse-papiers de Windows                          *
' *----------------------------------------------------------------------------------------------------------*
Option Explicit

Private Const CF_TEXT = 1               'Text format. Each line ends with a carriage return/linefeed (CR-LF) combination. A null character signals the end of the data. Use this format for ANSI text.
Private Const CF_BITMAP = 2             'A handle to a bitmap (HBITMAP).
Private Const CF_METAFILEPICT = 3       'Handle to a metafile picture format as defined by the METAFILEPICT structure. When passing a CF_METAFILEPICT handle by means of DDE, the application responsible for deleting hMem should also free the metafile referred to by the CF_METAFILEPICT handle.
Private Const CF_SYLK = 4               'Microsoft Symbolic Link (SYLK) format.
Private Const CF_TIFF = 6               'Tagged-image file format.
Private Const CF_DIF = 5                'Software Arts' Data Interchange Format.
Private Const CF_OEMTEXT = 7            'Text format containing characters in the OEM character set. Each line ends with a carriage return/linefeed (CR-LF) combination. A null character signals the end of the data.
Private Const CF_DIB = 8                'A memory object containing a BITMAPINFO structure followed by the bitmap bits.
Private Const CF_PALETTE = 9            'Handle to a color palette. Whenever an application places data in the clipboard that depends on or assumes a color palette, it should place the palette on the clipboard as well.
Private Const CF_PENDATA = 10           'Data for the pen extensions to the Microsoft Windows for Pen Computing.
Private Const CF_RIFF = 11              'Represents audio data more complex than can be represented in a CF_WAVE standard wave format.
Private Const CF_WAVE = 12              'Represents audio data in one of the standard wave formats, such as 11 kHz or 22 kHz PCM.
Private Const CF_UNICODETEXT = 13       'Unicode text format. Each line ends with a carriage return/linefeed (CR-LF) combination. A null character signals the end of the data.
Private Const CF_ENHMETAFILE = 14       'A handle to an enhanced metafile (HENHMETAFILE).
Private Const CF_HDROP = 15             'A handle to type HDROP that identifies a list of files. An application can retrieve information about the files by passing the handle to the DragQueryFile function.
Private Const CF_LOCALE = 16            'The data is a handle to the locale identifier associated with text in the clipboard. When you close the clipboard, if it contains CF_TEXT data but no CF_LOCALE data, the system automatically sets the CF_LOCALE format to the current input language. You can use the CF_LOCALE format to associate a different locale with the clipboard text.
Private Const CF_DIBV5 = 17             'A memory object containing a BITMAPV5HEADER structure followed by the bitmap color space information and the bitmap bits.
Private Const CF_DSPBITMAP = &H82       'Bitmap display format associated with a private format. The hMem parameter must be a handle to data that can be displayed in bitmap format in lieu of the privately formatted data.
Private Const CF_DSPENHMETAFILE = &H8E  'Enhanced metafile display format associated with a private format. The hMem parameter must be a handle to data that can be displayed in enhanced metafile format in lieu of the privately formatted data.
Private Const CF_DSPMETAFILEPICT = &H83 'Metafile-picture display format associated with a private format. The hMem parameter must be a handle to data that can be displayed in metafile-picture format in lieu of the privately formatted data.
Private Const CF_DSPTEXT = &H81         'Text display format associated with a private format. The hMem parameter must be a handle to data that can be displayed in text format in lieu of the privately formatted data.
Private Const CF_GDIOBJFIRST = &H300    'Start of a range of integer values for application-defined GDI object clipboard formats. The end of the range is CF_GDIOBJLAST.
Private Const CF_GDIOBJLAST = &H3FF     'See CF_GDIOBJFIRST.
Private Const CF_OWNERDISPLAY = &H80    'Owner-display format. The clipboard owner must display and update the clipboard viewer window, and receive the WM_ASKCBFORMATNAME, WM_HSCROLLCLIPBOARD, WM_PAINTCLIPBOARD, WM_SIZECLIPBOARD, and WM_VSCROLLCLIPBOARD messages. The hMem parameter must be NULL.

' Found 64-bit API declarations here: http://spreadsheet1.com/uploads/3/0/6/6/3066620/win32api_ptrsafe.txt
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
Private Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As LongPtr

' *---------------------------------------------------------------------------------------------------*
' * Copier une chaîne de caractères dans le presse-papiers de Windows                                 *
' *---------------------------------------------------------------------------------------------------*
Public Sub EcrirePressePapiers(sChaine As String)
    
    Const GMEM_MOVEABLE = &H2
    Const GMEM_ZEROINIT = &H40
    Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
    
    Dim hGlobalMemory As LongPtr, lpGlobalMemory As LongPtr
    
    ' Allouer de la mémoire globale
    hGlobalMemory = GlobalAlloc(GHND, Len(sChaine) + 1)
    ' Verrouiller la mémpoire afin d'obtnir un pointeur vers ce bloc
    lpGlobalMemory = GlobalLock(hGlobalMemory)
    ' Copier la chaîne de caractères vers la mémoire globale
    lpGlobalMemory = lstrcpy(lpGlobalMemory, sChaine)
    ' Déverrouille la méoire
    If GlobalUnlock(hGlobalMemory) <> 0 Then
        Err.Raise vbObjectError + 1, "Presse-papiers", "Erreur technique : impossible de déverrouiller l'emplacement mémoire."
    Else
        ' Ouvrir le presse-papiers afin de pouvoir copier
        If OpenClipboard(0&) = 0 Then
            Err.Raise vbObjectError + 2, "Presse-papiers", "Le presse-papiers ne peut pas être ouvert."
        Else
            ' Vider le presse-papiers
            Call EmptyClipboard
            ' Copier les données dans le presse-papiers
            Call SetClipboardData(CF_TEXT, hGlobalMemory)
            ' Fermer le presse-papiers (afin de valider la copie)
            If CloseClipboard() = 0 Then
                Err.Raise vbObjectError + 3, "Presse-papiers", "Le presse-papiers ne peut pas être fermé."
            End If
        End If
    End If
End Sub

' *---------------------------------------------------------------------------------------------------*
' * Lire le contenu du Presse-Papiers                                                                 *
' *---------------------------------------------------------------------------------------------------*
Public Function LirePressePapiers() As String

    ' Taille maximale des données texte pouvant être récupérées
    Const MAXSIZE = 1048576

    Dim hGlobalMemory As LongPtr, lpGlobalMemory As LongPtr, RetVal As LongPtr
    Dim sPressePapiers As String, lRet As Long
 
    LirePressePapiers = ""
    If OpenClipboard(0&) = 0 Then
        Err.Raise vbObjectError + 4, "Presse-papiers", "Le presse-papiers ne peut pas être ouvert."
    Else
        ' Obtenir le handle du bloc mémoire utilisé par le presse-papiers
        hGlobalMemory = GetClipboardData(CF_TEXT)
        If IsNull(hGlobalMemory) Then
            Err.Raise vbObjectError + 5, "Presse-papiers", "Erreur technique : impossible de récupérer l'emplacement mémoire."
        Else
            ' Verrouiller la mémoire du presse-papiers afin de pouvoir l'adresser
            lpGlobalMemory = GlobalLock(hGlobalMemory)
            ' Erreur lors du verrouillage ?
            If Not IsNull(lpGlobalMemory) Then
                ' Remplir la chaîne de caractères cible avec des caractères Null (0 binaire)
                sPressePapiers = Space$(MAXSIZE)
                Call lstrcpy(sPressePapiers, lpGlobalMemory)
                Call GlobalUnlock(hGlobalMemory)
                ' Rechercher le 1er caractère Null qui fait office de fin de chaîne
                If InStr(1, sPressePapiers, vbNullChar, vbBinaryCompare) > 0 Then
                    sPressePapiers = Mid(sPressePapiers, 1, InStr(1, sPressePapiers, vbNullChar, vbBinaryCompare) - 1)
                    LirePressePapiers = sPressePapiers
                End If
            Else
                Err.Raise vbObjectError + 6, "Presse-papiers", "Erreur technique : impossible de verrouiller l'emplacement mémoire."
            End If
        End If
        If CloseClipboard() = 0 Then
            Err.Raise vbObjectError + 3, "Presse-papiers", "Le presse-papiers ne peut pas être fermé."
        End If
    End If

End Function


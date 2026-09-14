Attribute VB_Name = "Ariana"
' ============================================================
'  Macro "Nueva OC"  -  Ariana / Oficina Tecnica
'  Al ejecutarlo: pide el CLIENTE y arma un mail nuevo con el
'  asunto y el cuerpo en formato estandar para que la
'  extraccion sea rapida y sin errores.
'
'  Para    :  orders@example.com  (destinatario, precargado)
'  CC      :  assistant@example.com  (siempre en copia)
'  Asunto  :  OC-<CLIENTE>-<fecha de hoy dd/mm/aaaa>
'  Cuerpo  :  N de OC del cliente:
'             MONTO SIN IVA:
'             Fecha de entrega:
'             Fabricacion (SI/NO):   -> SI = pendiente de fabricacion (GASPI)
'                                       NO o vacio = stock / estanteria
'             Proveedor:
'
'  CUENTA EMISORA:
'    Por defecto el mail sale desde la cuenta PREDETERMINADA de Outlook.
'    Si en esa PC la cuenta por defecto es otra (ej. la de Diego toma
'    la de Erik Solberg), completar CUENTA_EMISOR con el mail correcto
'    y el macro forzara ESA cuenta. Si se deja en "" usa la predeterminada.
'    (Si la cuenta indicada no existe en ese Outlook, usa la predeterminada.)
' ============================================================
Sub NuevaOC()
    ' >>> CONFIGURACION: dejar "" para usar la cuenta por defecto de Outlook,
    '     o poner el mail del vendedor para forzar el emisor (ej. Diego).
    Const CUENTA_EMISOR As String = "d.torres@example.com"

    Dim cliente As String
    cliente = Trim(InputBox("Ingresa el nombre del CLIENTE:", "Nueva OC - Ariana"))
    If cliente = "" Then Exit Sub          ' cancelado o vacio -> no hace nada
    cliente = UCase(cliente)

    Dim fechaHoy As String
    fechaHoy = Format(Date, "dd/mm/yyyy")  ' fecha de hoy, siempre actual

    Dim nOC As String
    nOC = "N" & Chr(176) & " de OC del cliente:"   ' Chr(176) = simbolo "°"

    Dim olMail As Object
    Set olMail = Application.CreateItem(0)   ' 0 = olMailItem

    ' --- Forzar la cuenta emisora si esta configurada y existe en este Outlook ---
    If CUENTA_EMISOR <> "" Then
        Dim acc As Object
        For Each acc In Application.Session.Accounts
            If LCase(acc.SmtpAddress) = LCase(CUENTA_EMISOR) Then
                Set olMail.SendUsingAccount = acc
                Exit For
            End If
        Next
    End If
    ' ---------------------------------------------------------------------------

    olMail.To = "orders@example.com"              ' destinatario precargado
    olMail.CC = "assistant@example.com"     ' siempre en copia
    olMail.Subject = "OC-" & cliente & "-" & fechaHoy
    olMail.Body = _
        nOC & " " & vbCrLf & _
        "MONTO SIN IVA: " & vbCrLf & _
        "Fecha de entrega: " & vbCrLf & _
        "Fabricacion (SI/NO): " & vbCrLf & _
        "Proveedor: " & vbCrLf

    olMail.Display   ' abre el mail ya armado para completar y enviar
End Sub

<# 

Desarrollado por: Marcos A. Nunes
Fecha: 08/08/2023
Versión: 3.6.6 Bogota
Aplicación: Script Helper

Descripción: Esta aplicación es una herramienta de soporte para automatización de tareas en scripts de PowerShell, 
proporcionando funciones y recursos para gestionar usuarios, cuentas, información y otras operaciones específicas del entorno de Active Directory.
Originalmente desarrollado para el Service Desk Campo Grande, para la cuenta Heineken.

Traducción realizada a través de un traductor online; en caso de problemas, por favor, póngase en contacto para corrección.

#> 

# Importando o módulo ActiveDirectory
# Carregar objetos para exibição de MsgBox
Import-Module ActiveDirectory
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.DirectoryServices.AccountManagement

# Função para exibir o menu e obter a opção escolhida pelo usuário
function Show-Menu {
    Write-Host "========== Menú =========="
    Write-Host "1. Desbloquear cuenta"
    Write-Host "2. Activar/Desactivar cuenta"
    Write-Host "3. Imprimir grupos"
    Write-Host "4. Restablecer contraseña"
    Write-Host "5. Buscar Gerente"
    Write-Host "6. Cambiar teléfono"
    Write-Host "7. Copiar información"
    Write-Host "0. Salir"
    Write-Host "==========================="
    return Read-Host "Elija una opción"
}

# Função para desbloquear a conta
function Unlock-Account($account) {
    if ($account.LockedOut) {
        Write-Host "La cuenta de usuario está bloqueada." -ForegroundColor Red
        Write-Host "Escriba 'd' para desbloquear la cuenta o presione Enter para continuar:" -NoNewline -ForegroundColor Red

        $choice = Read-Host 

        if ($choice -ieq 'd') {
            # Desbloquear a conta do usuário
            Unlock-ADAccount -Identity $account.DistinguishedName -Credential $Credential
            Write-Host "La cuenta de usuario ha sido desbloqueada." -ForegroundColor Green
        }
    }
    else {
        Write-Host "La cuenta de usuario no está bloqueada." -ForegroundColor Green
        Write-Host "Escriba 'd' para desbloquearla forzadamente o presione Enter para continuar:" -NoNewline -ForegroundColor Red
        $choice = Read-Host

        if ($choice -ieq 'd') {
            # Desbloquear a conta do usuário
            Unlock-ADAccount -Identity $account.DistinguishedName -Credential $Credential
            Write-Host "La cuenta del usuario ha sido desbloqueada." -ForegroundColor Green
        } 
    }
}

# Função para ativar ou desativar a conta
function Enable-Disable-Account($account) {
    if ($account.Enabled) {
        Write-Host "La cuenta del usuario está ACTIVA." -ForegroundColor Green
        Write-Host "¿Está seguro de que desea desactivarla? Presione 'S' para DESACTIVAR o Enter para continuar:" -NoNewline -ForegroundColor Red
        $choice = Read-Host 

        if ($choice -ieq 'S') {
            # Desativar a conta do usuário
            $account | Disable-ADAccount -Credential $Credential
            Write-Host "La cuenta ha sido DESACTIVADA." -ForegroundColor Red
        }
        else {
            Write-Host "La cuenta del usuario sigue ACTIVA" -ForegroundColor Green
        }

    }
    else {
        Write-Host "La cuenta del usuario está DESACTIVADA." -ForegroundColor Red
        Write-Host "¿Está seguro de que desea ACTIVAR? Presione 'S' para ACTIVAR o Enter para continuar:" -NoNewline -ForegroundColor Yellow

        $choice = Read-Host 

        if ($choice -ieq 'S') {
        #Ativar conta no AD
        $account | Enable-ADAccount -Credential $Credential
        Write-Host "La cuenta ha sido ACTIVADA." -ForegroundColor Green
        }
        else {
            Write-Host "La cuenta del usuario sigue DESACTIVADA" -ForegroundColor Red
        }
    }
}

# Mostra os grupos da conta
function Show-Groups($account) {
    $GroupNames = (Get-ADUser $account.SamAccountName -Properties MemberOf).MemberOf | Get-ADGroup | Select-Object -ExpandProperty Name | Sort-Object

    if ($GroupNames.Count -eq 0) {
        Write-Host "El usuario no pertenece a ningún grupo." -ForegroundColor Red
        return
    }

    Add-Type -AssemblyName System.Windows.Forms
    $Form = New-Object System.Windows.Forms.Form
    $ListBox = New-Object System.Windows.Forms.ListBox
    $TextBoxSearch = New-Object System.Windows.Forms.TextBox
    $ButtonSearch = New-Object System.Windows.Forms.Button

    $Form.Size = New-Object System.Drawing.Size(430, 500)
    $Form.Controls.Add($ListBox)
    $Form.Controls.Add($TextBoxSearch)
    $Form.Controls.Add($ButtonSearch)

    $ListBox.Size = New-Object System.Drawing.Size(380, 350)
    $ListBox.Location = New-Object System.Drawing.Point(10, 10)
    $ListBox.Anchor = 'Top, Left, Right, Bottom'

    $TextBoxSearch.Size = New-Object System.Drawing.Size(300, 25)
    $TextBoxSearch.Location = New-Object System.Drawing.Point(10, 370)
    $TextBoxSearch.Anchor = 'Left, Bottom'
    $TextBoxSearch.Add_KeyDown({
            if ($_.KeyCode -eq 'Enter') {
                $ButtonSearch.PerformClick()
            }
        })

    $ButtonSearch.Size = New-Object System.Drawing.Size(80, 25)
    $ButtonSearch.Text = "Buscar"
    $ButtonSearch.Location = New-Object System.Drawing.Point(320, 370)
    $ButtonSearch.Anchor = 'Right, Bottom'
    $ButtonSearch.Add_Click({
            $searchTerm = $TextBoxSearch.Text
            if ($searchTerm -ne '') {
                $filteredGroups = $GroupNames | Where-Object { $_ -like "*$searchTerm*" }
                if ($filteredGroups.Count -gt 0) {
                    $ListBox.Items.Clear()
                    $ListBox.Items.AddRange($filteredGroups)
                }
                else {
                    [System.Windows.Forms.MessageBox]::Show("Ningún grupo encontrado para el término de búsqueda.: $searchTerm", "Búsqueda de Grupos", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                }
            }
            else {
                $ListBox.Items.Clear()
                $ListBox.Items.AddRange($GroupNames)
            }
        })

    $ListBox.Items.AddRange($GroupNames)

    $ListBox.Add_DoubleClick({
            $selectedGroup = $ListBox.SelectedItem
            [System.Windows.Forms.Clipboard]::SetText($selectedGroup)
        })

    $Form.ShowDialog() | Out-Null
}



# Função para redefinir a senha
# Cria caixa de dialogo para escrever a senha ou colar
function Show-CustomCredentialDialog($message, $username = "") {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Ingrese la nueva contraseña"
    $form.Width = 300
    $form.Height = 200
    $form.StartPosition = "CenterScreen"

    # Exibe a mensagem na caixa de diálogo
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 20)
    $label.Size = New-Object System.Drawing.Size(280, 20)
    $label.Text = $message
    $form.Controls.Add($label)

    # Exibe o user na caixa de diálogo
    $textbox = New-Object System.Windows.Forms.Label
    $textbox.Location = New-Object System.Drawing.Point(10, 50)
    $textbox.Size = New-Object System.Drawing.Size(280, 20)
    $textbox.Text = $username.name
    $form.Controls.Add($textbox)

    # Recebe a senha
    $passwordbox = New-Object System.Windows.Forms.TextBox
    $passwordbox.Location = New-Object System.Drawing.Point(10, 80)
    $passwordbox.Size = New-Object System.Drawing.Size(280, 20)
    $passwordbox.UseSystemPasswordChar = $false
    $form.Controls.Add($passwordbox)

    # Botão de OK para confirmar a ação
    $button = New-Object System.Windows.Forms.Button
    $button.Location = New-Object System.Drawing.Point(100, 120)
    $button.Size = New-Object System.Drawing.Size(100, 30)
    $button.Text = "OK"
    $button.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($button)

    $form.AcceptButton = $button

    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        #Valida se a senha esta vazia ou nula
        if ([string]::IsNullOrEmpty($passwordbox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Proporcione una contraseña válida.", "Erro", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        $password = $passwordbox.Text | ConvertTo-SecureString -AsPlainText -Force
        $credential = New-Object System.Management.Automation.PSCredential -ArgumentList $username.name, $password
        return $credential
    }
}
# Função para redefinir senha
function Reset-Password($account) {
    $maxAttempts = 2
    $attemptCount = 0

    do {
        try {
            $cred = Show-CustomCredentialDialog -message "Ingrese la nueva contraseña" -username $account
            $newPassword = $cred.GetNetworkCredential().Password | ConvertTo-SecureString -AsPlainText -Force

            Set-ADAccountPassword -Identity $account -NewPassword $newPassword -Credential $Credential -ErrorAction Stop

            $changePasswordAtLogon = Read-Host -Prompt "¿Desea que la contraseña sea cambiada al iniciar sesión? (S/N)"

            if ($changePasswordAtLogon -eq "S" -or $changePasswordAtLogon -eq "s") {
                Set-ADUser -Identity $account -ChangePasswordAtLogon $true -Credential $Credential
                Write-Host "¡La contraseña del usuario ha sido restablecida exitosamente! La contraseña será cambiada al iniciar sesión." -ForegroundColor Green
            }
            else {
                Write-Host "¡La contraseña del usuario ha sido restablecida exitosamente! La contraseña no será cambiada al iniciar sesión." -ForegroundColor Green
            }

            $validPassword = $true
        }
        catch {
            Write-Host "Ha ocurrido un error al restablecer la contraseña del usuario." -ForegroundColor Red
            Write-Host $_.Exception.Message -ForegroundColor Red

            $attemptCount++

            if ($attemptCount -eq 1) {
                Write-Host "Ejemplo de contraseña válida: Brasil@******" -ForegroundColor Cyan
            }

            $validPassword = $false
        }
    } while (-not $validPassword -and $attemptCount -lt $maxAttempts)
}



# Redefine o telefone do usuário, pode alternar entre mobile e telephone
function Set-ADUserPhoneNumber {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Username
    )

    try {
        # Abre a caixa de diálogo para inserir o número de telefone
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "Cambiar número de teléfono"
        $form.Width = 320
        $form.Height = 250
        $form.StartPosition = "CenterScreen"

        # Exibe o nome de usuário na caixa de diálogo
        $label = New-Object System.Windows.Forms.Label
        $label.Location = New-Object System.Drawing.Point(10, 20)
        $label.Size = New-Object System.Drawing.Size(280, 20)
        $label.Text = "Escriba el número de teléfono para el usuario:"        
        $form.Controls.Add($label)

        # Recebe o número de telefone
        $phoneNumberBox = New-Object System.Windows.Forms.TextBox
        $phoneNumberBox.Location = New-Object System.Drawing.Point(10, 50)
        $phoneNumberBox.Size = New-Object System.Drawing.Size(280, 20)
        $form.Controls.Add($phoneNumberBox)

        # Caixa de seleção para escolher o tipo de telefone
        $phoneTypeComboBox = New-Object System.Windows.Forms.ComboBox
        $phoneTypeComboBox.Location = New-Object System.Drawing.Point(10, 80)
        $phoneTypeComboBox.Size = New-Object System.Drawing.Size(280, 20)
        $phoneTypeComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
        $phoneTypeComboBox.Items.AddRange(@("Mobile", "Telephone"))
        $form.Controls.Add($phoneTypeComboBox)

        # Botão de OK para confirmar a ação
        $button = New-Object System.Windows.Forms.Button
        $button.Location = New-Object System.Drawing.Point(100, 120)
        $button.Size = New-Object System.Drawing.Size(100, 30)
        $button.Text = "OK"
        $button.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Controls.Add($button)

        $form.AcceptButton = $button

        $result = $form.ShowDialog()

        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            $phoneNumber = $phoneNumberBox.Text.Trim()
            $phoneType = $phoneTypeComboBox.SelectedItem.ToString()

            if ([string]::IsNullOrWhiteSpace($phoneNumber)) {
                Write-Host "Por favor, proporcione un número de teléfono válido." -ForegroundColor Red
                return
            }

            # Atualiza o número de telefone no AD
            $user = Get-ADUser -Identity $Username -ErrorAction Stop

            switch ($phoneType) {
                "Mobile" {
                    $user.MobilePhone = $phoneNumber
                }
                "Telephone" {
                    $user.TelephoneNumber = $phoneNumber
                }
            }
            # Faz alteração do telefone no AD
            Set-ADUser -Instance $user -ErrorAction Stop

            Write-Host "El número de teléfono se ha actualizado con éxito para el usuario." -ForegroundColor Green
        }
        else {
            Write-Host "Operación cancelada." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "Ha ocurrido un error al actualizar el número de teléfono del usuario. '$Username': $($_.Exception.Message)" -ForegroundColor Red
    }
}


# Função para pegar as informações do gestor
function Get-ManagerInfo($User) {
    if ($User.Manager) {
        $Manager = Get-ADUser -Identity $User.Manager -Properties * 
        
        $InfoManager = @"
• Nombre: $($Manager.DisplayName)
• Heiway: $($Manager.Name)
• Matrícula: $($Manager.extensionAttribute5)
• Correo electrónico: $($Manager.EmailAddress)
• Localidad: $($Manager.Office)
• Departamento: $($Manager.Department)
• Puesto: $($Manager.Title)
• Teléfono: $($Manager.MobilePhone)
"@

        # Copiar informações para a área de transferência
        [System.Windows.Forms.Clipboard]::SetText($InfoManager)
        # Exibe as informações do gestor no console
        Write-Host "$InfoManager`nLas informaciones del gestor han sido copiadas." -ForegroundColor Green

    }
    else {
        Write-Host "No hay un gestor definido para el usuario." -ForegroundColor Red
    }
}

# Configurar o console
function ConfigConsole {
    [CmdletBinding()]
    param()

    try {


        # Configuração da aparência do PowerShell
        $Host.UI.RawUI.BackgroundColor = "Black"  # Cor do fundo
   

        # Definir a cor do texto
        $Host.UI.RawUI.ForegroundColor = "White"  # Cor do texto 
        

        # Configuração do tamanho do console
        $Width = 92
        $Height = 40

        $WindowSize = New-Object System.Management.Automation.Host.Size ($Width, $Height)
        $Host.UI.RawUI.WindowSize = $WindowSize

        # Obter o tamanho total da tela
        $ScreenSize = $Host.UI.RawUI.WindowSize
        $ScreenWidth = $ScreenSize.Width

        # Calcular a posição centralizada
        $Left = [math]::Max(0, ($ScreenWidth - $Width) / 2)
        $Top = 0

        # Definir a nova posição do console
        $NewPosition = New-Object System.Management.Automation.Host.Coordinates -ArgumentList ($Left, $Top)
        $Host.UI.RawUI.WindowPosition = $NewPosition

        Write-Host "¡Consola de PowerShell configurada exitosamente!" -ForegroundColor Green
    }
    catch {
        Write-Host "Ocurrió un error durante la configuración de la consola.: $_" -ForegroundColor Red
    }
}

# Função para copiar os dados da conta
function Get-AccountInfo {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [string]$UserName
    )

    $AttributesToRetrieve = @(
        "DisplayName",
        "SamAccountName",
        "Name",
        "extensionAttribute5",
        "EmailAddress",
        "Office",
        "Department",
        "Title",
        "TelephoneNumber",
        "MobilePhone",
        "HomePhone",
        "OtherTelephone",
        "ipPhone"
    )

    $User = Get-ADUser -Filter { SamAccountName -eq $UserName -or extensionAttribute5 -eq $UserName -or EmailAddress -eq $UserName -or DisplayName -eq $UserName } -Properties $AttributesToRetrieve

    $InfoAccount = @"
• Nombre: $($User.DisplayName)
• Heiway: $($User.Name)
• Matrícula: $($User.extensionAttribute5)
• Correo electrónico: $($User.EmailAddress)
• Localidad: $($User.Office)
• Departamento: $($User.Department)
• Función: $($User.Title)
"@

    # Filtrar apenas os atributos relacionados aos telefones
    $PhoneAttributes = $AttributesToRetrieve | Where-Object { $_ -match 'TelephoneNumber|MobilePhone|HomePhone|OtherTelephone|ipPhone' }

    # Contador de telefones
    $PhoneCount = 0

    # Percorrer cada atributo de telefone e adicionar o valor à string se estiver definido
    foreach ($Attribute in $PhoneAttributes) {
        if ($User.$Attribute) {
            $AttributeValue = $User.$Attribute
            $InfoAccount += "`n• $Attribute"
            $InfoAccount += ": $AttributeValue"

            $PhoneCount++
        }
    }

    if ($PhoneCount -eq 0) {
        $InfoAccount += "`n• Teléfono: "
    }

    # Copiar informações para a área de transferência
    Add-Type -AssemblyName "System.Windows.Forms"
    [System.Windows.Forms.Clipboard]::SetText($InfoAccount)

    Write-Host "Las informaciónes han sido copiadas." -ForegroundColor Green
}

# Limpar console
function Clear-Console {
    Clear-Host
}

# Exibir Cabeçalho 
function Show-Header {
    $combinedLetters = @"
                       

 _______ _      _        _      _    _      _                 
|__   __(_)    | |      | |    | |  | |    | |                
   | |   _  ___| | _____| |_   | |__| | ___| |_ __   ___ _ __ 
   | |  | |/ __| |/ / _ \ __|  |  __  |/ _ \ | '_ \ / _ \ '__|
   | |  | | (__|   <  __/ |_   | |  | |  __/ | |_) |  __/ |   
   |_|  |_|\___|_|\_\___|\__|  |_|  |_|\___|_| .__/ \___|_|   
   Service Desk Bogotá                       | |              
                                             |_|              

"@

    Write-Host "==============================================================================" -ForegroundColor Yellow   
    # Write-Host $star -ForegroundColor Red
    Write-Host $combinedLetters -ForegroundColor Red              
    Write-Host "==============================================================================" -ForegroundColor Yellow
}
# Pega informações do usuário no AD
function update-user {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SearchQuery
    )

    try {
        # Realiza a busca com base no atributo SamAccountName, extensionAttribute5, EmailAddress e DisplayName
        Return Get-ADUser -Filter { SamAccountName -eq $SearchQuery -or extensionAttribute5 -eq $SearchQuery -or EmailAddress -eq $SearchQuery -or DisplayName -eq $SearchQuery } -Properties *

                
    }
    catch {
        Write-Host "Lo siento, ocurrió un error durante la ejecución de la función:"
        Write-Host $_.Exception.Message
    }
}




# Informações do usuário, nessa função são realizados testes e exibe no console de acordo.
function Show-UserInfo {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [string]$UserName
    )

    try {
        # Chama a função para atualizar o usuário
        $User = update-user $UserName

        if ($User.extensionAttribute8 -eq "VIP") {
            $ForegroundColor = "Cyan"
            Write-Host "==================== *** USUÁRIO VIP *** ====================" -ForegroundColor $ForegroundColor
        }
        else {
            $ForegroundColor = "Green"
        }

        Write-Host "==================== Informaciones del Usuario ===================="
        Write-Host "• Nombre: $($User.DisplayName)" -ForegroundColor $ForegroundColor
        Write-Host "• Heiway: $($User.Name)" -ForegroundColor $ForegroundColor
        Write-Host "• Matrícula: $($User.extensionAttribute5)" -ForegroundColor $ForegroundColor
        Write-Host "• Correo electrónico: $($User.EmailAddress)" -ForegroundColor $ForegroundColor
        Write-Host "• Localidad: $($User.Office)" -ForegroundColor $ForegroundColor
        Write-Host "• Departamento: $($User.Department)" -ForegroundColor $ForegroundColor
        Write-Host "• Función: $($User.Title)" -ForegroundColor $ForegroundColor

        
        # Verifica todos os atributos onde pode conter um telefone e os imprime, caso tenha algum valor
        $PhoneAttributes = @(
            "TelephoneNumber",
            "MobilePhone",
            "HomePhone",
            "OtherTelephone",
            "ipPhone"
        )

        # Percorrer cada atributo e imprimir o valor se estiver definido
        foreach ($Attribute in $PhoneAttributes) {
            if ($User.$Attribute) {
                $AttributeValue = $User.$Attribute
                Write-Host "• $Attribute" -NoNewline -ForegroundColor $ForegroundColor
                Write-Host ": $AttributeValue" -ForegroundColor $ForegroundColor

            }
        }

        # Exibe informações básicas do gestor
        if ($User.Manager) {
            $Manager = Get-ADUser -Identity $User.Manager -Properties DisplayName, Name
            $ManagerDisplayName = $Manager.DisplayName
            $ManagerUserName = $Manager.Name
            Write-Host "• Gerente: $ManagerDisplayName, $ManagerUserName" -ForegroundColor $ForegroundColor
        }

        Write-Host "----------------------------- INFO -----------------------------"
        Write-Host "• Tipo de empleado: $($User.employeeType)" -ForegroundColor $ForegroundColor

        # Verifica o valor do attribute11 e define a cor do texto
        if ($User.extensionAttribute11 -match "DESKLESS") { 
            $ForegroundColor = "Yellow" 
        }
        elseif ($User.extensionAttribute11 -match "ENTERPRISE") { 
            $ForegroundColor = "Green" 
        }
        else { 
            $ForegroundColor = "Red" 
        }

        Write-Host "• Tipo de licencia de Office: $($User.extensionAttribute11)" -ForegroundColor $ForegroundColor

        # Volta à cor original
        $ForegroundColor = "Green"

        Write-Host "• Tipo de proxy: $($User.extensionAttribute7)" -ForegroundColor $ForegroundColor
        Write-Host "---------------------------- STATUS ----------------------------"

        # Verifica o status da conta do usuário
        $StatusPositive = @()
        $StatusNegative = @()

        if ($User.Enabled) {
            $StatusPositive += " [+] Activo"
        }
        else {
            $StatusNegative += " [-] Desactivada"
            
        }

        # Verifica se a conta do usuário está bloqueada
        if ($User.LockedOut) {
            $StatusNegative += " [-] Bloqueada"
                    
        }
        else {
            $StatusPositive += " [+] No está bloqueada"
        }

        # Verifica se o usuário está de férias
        $GroupName = "BR1-AccountsHoliday-US"
        if ((Get-ADUser -Identity $User -Properties MemberOf).MemberOf -like "*$GroupName*") {
            $StatusNegative += " [-] Está de vacaciones"
            
        }
        else {
            $StatusPositive += " [+] No está de vacaciones"
        }

        # Impressão dos status positivos
        if ($StatusPositive.Count -gt 0) {
            Write-Host "• Cuenta de usuario:" $StatusPositive -ForegroundColor Green
        }

        # Impressão dos status negativos
        if ($StatusNegative.Count -gt 0) {
            Write-Host "• Cuenta de usuario:" $StatusNegative -ForegroundColor Red
            Write-Host "Descripción: $($User.Description)" -ForegroundColor Cyan
        }

        Write-Host "----------------------------------------------------------------"
         
        # Verifica a data de expiração da conta
        $accountExpires = [datetime]::FromFileTime($User.accountExpires)   
        $accountExpiresString = $accountExpires.ToString("dd/MM/yyyy HH:mm:ss")
        $currentDate = Get-Date
        
        if ($null -ne $accountExpiresString) {
            if ($accountExpires -lt $currentDate) {
                Write-Host "• Status da cuenta: [-] Caducada" -ForegroundColor Red
                Write-Host "Descripción: $($User.Description)" -ForegroundColor Cyan
            }
            elseif ($accountExpires -gt $currentDate) {
                
                Write-Host "• Status da cuenta: [+] No Caducada" -ForegroundColor Green
                # Verifica quantos dias faltam para a conta expirar
                $daysUntilAccountExpiration = ($accountExpires - $currentDate).Days
                            
            }
            else {
                Write-Host "• Status da cuenta: [!] Caduca hoje" -ForegroundColor Yellow
            }

            Write-Host "Cuenta Caduca em: $accountExpiresString" - "Faltam: $daysUntilAccountExpiration dias" -ForegroundColor Cyan
        }
        else {
            Write-Host "• Status da cuenta: [+] Sin fecha de vencimiento definida" -ForegroundColor Green
        }

        Write-Host "---------------------------- Contraseña -----------------------------"

        # Verifica se a senha do usuário será alterada no próximo logon
        if ($null -eq $User.PasswordLastSet) {
            Write-Host "• Cambio de contraseña en el próximo inicio de sesión: [!] Sí" -ForegroundColor Yellow

        }
        else {
            Write-Host "• Cambio de contraseña en el próximo inicio de sesión: [+] Não" -ForegroundColor Green

            # Verifica a validade da senha (1 ano após ter sido definida)
            $passwordExpirationDate = $User.PasswordLastSet.AddDays(365)
            if ($passwordExpirationDate -lt $currentDate) {
                Write-Host "• Validez de la contraseña: [-] Caducada" -ForegroundColor Red
            }
            else {
                Write-Host "• Validez de la contraseña: [+] No caducada" -ForegroundColor Green
                $daysUntilExpiration = ($passwordExpirationDate - $currentDate).Days
                
            }
            
            Write-Host "Contraseña establecida el: $($User.PasswordLastSet.ToString("dd/MM/yyyy HH:mm:ss")) - Faltam: $daysUntilExpiration dias" -ForegroundColor Cyan

        }

        Write-Host "----------------------------------------------------------------"

    }
    catch {
        # Tratamento para a falha com data: "Not a valid Win32 FileTime"
        if ($_.Exception.Message -like "*Not a valid Win32 FileTime*") {
                
                
            Write-Host "---------------------------- Contraseña -----------------------------"

            # Verifica se a senha do usuário será alterada no próximo logon
            if ($null -eq $User.PasswordLastSet) {
                Write-Host "• Cambio de contraseña en el próximo inicio de sesión: [!] Sí" -ForegroundColor Yellow

            }
            else {
                Write-Host "• Cambio de contraseña en el próximo inicio de sesión: [+] Não" -ForegroundColor Green

                # Verifica a validade da senha (1 ano após ter sido definida)
                $passwordExpirationDate = $User.PasswordLastSet.AddDays(365)
                if ($passwordExpirationDate -lt $currentDate) {
                    Write-Host "• Validez de la contraseña: [-] Caducada" -ForegroundColor Red
                }
                else {
                    Write-Host "• Validez de la contraseña: [+] No caducada" -ForegroundColor Green                
                
                }

                Write-Host "No se pudo cargar el contador de días." -ForeGroundColor Red
            
            }

            Write-Host "----------------------------------------------------------------"

        }
        else {
            Write-Host "Se produjo un error al obtener la información del usuario." -ForeGroundColor Red
            Write-Host $_.Exception.Message
        }
   
    }
}
#Execução antes do loop Principal
Clear-Console
Write-Host @' 

Desarrollado por: Marcos A. Nunes
Fecha: 08/08/2023
Versión: 3.6.6 Bogota
Aplicación: Script Helper

Descripción: Esta aplicación es una herramienta de soporte para automatización de tareas en scripts de PowerShell, 
proporcionando funciones y recursos para gestionar usuarios, cuentas, información y otras operaciones específicas 
del entorno de Active Directory.
Originalmente desarrollado para el Service Desk Campo Grande, para la cuenta Heineken.

Traducción realizada a través de un traductor online; en caso de problemas, por favor, póngase en contacto para corrección.

'@ -ForeGround Cyan

# Solicitar as credenciais do administrador
# Definir o número máximo de tentativas
$maxAttempts = 2
$attempts = 0

# Função teste de autenticação
Function Test-ADAuthentication {
    Param(
        [Parameter(Mandatory)]
        [string]$User,
        [Parameter(Mandatory)]
        $Password,
        [Parameter(Mandatory = $false)]
        $Server,
        [Parameter(Mandatory = $false)]
        [string]$Domain = $Env:USERDOMAIN
    )
    
    $contextT = [System.DirectoryServices.AccountManagement.ContextType]::Domain
    $contextO = [System.DirectoryServices.AccountManagement.ContextOptions]::Negotiate
    
    $argumentList = New-Object -TypeName "System.Collections.ArrayList"
    $null = $argumentList.Add($contextT)
    $null = $argumentList.Add($Domain)
    If ($null -ne $Server) {
        $argumentList.Add($Server)
    }
    
    $principalC = New-Object System.DirectoryServices.AccountManagement.PrincipalContext -ArgumentList $argumentList -ErrorAction SilentlyContinue
    if ($null -eq $principalC) {
        return $false
    }
    
    if ($principalC.ValidateCredentials($User, $Password, $contextO)) {
        return $true
    }
    else {
        return $false
    }
} 

# Uso da Função
Do {
    # Solicita usuário e senha de ADMINISTRADOR
    $Credential = Get-Credential -Message "Ingrese nombre de usuario y contraseña de administrador para validar el acceso a HEIWAY:"
    
    # Verifica se o usuário pressionou "Cancel"
    if ($null -eq $Credential) {
        [Microsoft.VisualBasic.Interaction]::MsgBox("Autenticación cancelada, algunas funciones no están disponibles.", 'OKOnly,SystemModal,Exclamation', 'Test-ADAuthentication')
        break
    }
    
    $User = $Credential | Select-Object Username -ExpandProperty Username
    $Pass = $Credential.GetNetworkCredential() | Select-Object Password -ExpandProperty Password
    # Retire o comentário da próxima linha para consultar outro domínio
    #$Domain = "outrodominio.local"
    
    # Executa função (Retire o comentário na próxima linha caso esteja utilizando outro domínio)
    $Validar = Test-ADAuthentication -User $User -Password $Pass #-Domain $Domain
 
    If ($null -eq $credential -or $Validar -eq $false) {
        # Incrementar o contador de tentativas
        $attempts++
        [Microsoft.VisualBasic.Interaction]::MsgBox("ERROR: Las credenciales para el usuario " + $User + " están incorrectas! Por favor, ingréselas nuevamente..", 'OKOnly,SystemModal,Exclamation', 'Test-ADAuthentication')
    }
    Else {
        [Microsoft.VisualBasic.Interaction]::MsgBox("ÉXITO: Las credenciales para el usuario están correctas. " + $User + " están correctas.!", 'OKOnly,SystemModal,Exclamation', 'Test-ADAuthentication')
    }

    # Verificar se o número máximo de tentativas foi atingido
    if ($attempts -ge $maxAttempts) {
        [Microsoft.VisualBasic.Interaction]::MsgBox("Se ha alcanzado el número máximo de intentos. Las credenciales son inválidas y algunas funciones del programa no funcionarán.", 'OKOnly,SystemModal,Exclamation', 'Test-ADAuthentication')
        break
    }
} While ($Validar -ne $true)

# Loop principal
while ($true) {
    
    ConfigConsole
    Clear-Console
    Show-Header

    $UserName = Read-Host "Ingrese el usuario (o escriba 'sair' para finalizar)"

    if ($UserName -eq "sair") {
        break
    }

    try {
        Clear-Console        
        $UserTest = Get-ADUser -Filter { SamAccountName -eq $UserName -or extensionAttribute5 -eq $UserName -or EmailAddress -eq $UserName -or DisplayName -eq $UserName }

        if ($null -eq $UserTest) {
            [Microsoft.VisualBasic.Interaction]::MsgBox("Usuário não encontrado.", 'OKOnly,SystemModal,Exclamation', 'Test-ADAuthentication')
            continue
        }

        Clear-Console
        Show-UserInfo $UserName
        Get-AccountInfo $UserName        

        $exitLoop = $false
        while (-not $exitLoop) {
            $User = update-user $UserName        
            $option = Show-Menu
            switch ($option) {
                # Desbloqueio da conta
                1 {
                    Clear-Console
                    Show-UserInfo $UserName
                    Unlock-Account $User -Credential $Credential
                    
                }
                # Ativar e desativar a conta
                2 {
                    Clear-Console
                    Show-UserInfo $UserName
                    Enable-Disable-Account $User -Credential $Credential
                }
                # Printar os grupos
                3 {
                    Clear-Console
                    Show-UserInfo $UserName
                    Show-Groups $User
                }
                # Reset de senha de rede
                4 {
                    Clear-Console
                    Show-UserInfo $UserName
                    Reset-Password $User -Credential $Credential
                }
                # Buscar o Gestor
                5 {
                    Clear-Console
                    Show-UserInfo $UserName
                    Get-ManagerInfo $User
                }
                # Alterar o telefone
                6 {
                    Clear-Console
                    Show-UserInfo $UserName
                    Set-ADUserPhoneNumber $User
                }
                7 {
                    Clear-Console
                    Show-UserInfo $UserName
                    Get-AccountInfo $UserName
                }

                #Sair
                0 {
                    $exitLoop = $true
                }

                default {
                    Clear-Console
                    Show-UserInfo $UserName
                    Write-Host "Opción inválida. Intente nuevamente."
                }
            } 
        }
    }
    catch {
        Write-Host "Se produjo un error al obtener la información del usuario"
        Write-Host $_.Exception.Message
    }
}
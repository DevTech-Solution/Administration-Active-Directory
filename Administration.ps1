#######################################################################################################################################################################
                                                                     # OUTILS D'ADMINISTRATION par SEB #
#######################################################################################################################################################################
                                                                              # VIDER CONSOLE #
#######################################################################################################################################################################
Clear-Host
#######################################################################################################################################################################
                                                                        # Chargement des Windows Form #
#######################################################################################################################################################################
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void][Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')  
#######################################################################################################################################################################
                                                                         # CACHER FENETRE POWERSHELL #
#######################################################################################################################################################################
PowerShell.exe -WindowStyle Hidden -File 
#######################################################################################################################################################################
                                                                              # VIDER CONSOLE #
#######################################################################################################################################################################
Clear-Host
#######################################################################################################################################################################
###
 ###Ajouter votre domaine
$Domaine = "mondomaine.ad"
$UO = "OU=Instituts-Ecoles-Departements,DC=afpicl,DC=lan" 
$Server = "ad.enterprise"
#######################################################################################################################################################################
                                                                              # FENETRE MENU #
#######################################################################################################################################################################
#region FENETRE MENU
#MENU
$form = New-Object Windows.Forms.Form
$form.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $form.Close() } }) 
$form.KeyPreview = $true
$form.Height = 650
$form.Width = 400
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = 'Fixed3D' 
$form.Text = "OUTILS D'ADMINISTRATION"
$form.BackColor = "Lightblue"

#Label (TITRE 1)
Add-Type -AssemblyName System.Windows.Forms
$Label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Size(54,50)
$Label.Text = "OUTILS D'ADMINISTRATION"
$Label.Font = New-Object System.Drawing.Font("Times New Roman",15,[System.Drawing.FontStyle]::Bold)
$Label.AutoSize = $false
$Label.Size = New-Object System.Drawing.Size(320,30)
$Form.Controls.Add($Label)


#Label2 (Signature)
Add-Type -AssemblyName System.Windows.Forms
$Form.Font = $Font
$Label2 = New-Object System.Windows.Forms.Label
$label2.Location = New-Object System.Drawing.Size(165,590)
$Label2.Text = "SebCorp"
$Label2.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
$Label2.AutoSize = $false
$Label2.Size = New-Object System.Drawing.Size(70,60)
$Form.Controls.Add($Label2)
#endregion
#######################################################################################################################################################################
                                                                                # LABEL + LISTBOX (MENU) #
#######################################################################################################################################################################
#region LABEL+LISTBOX

#Label
$FormLabel_1 = New-Object System.Windows.Forms.Label
$FormLabel_1.Location = New-Object System.Drawing.Point(46,112)
$FormLabel_1.Size = New-Object System.Drawing.Size(280,15)
$FormLabel_1.Font = New-Object System.Drawing.Font("Arial",10,[System.Drawing.FontStyle]::Bold)
$FormLabel_1.Text = "Séléctionner votre choix :"
$form.Controls.Add($Formlabel_1) 

#Listbox
$ListBox_1 = New-Object System.Windows.Forms.ListBox 
$ListBox_1.Location = New-Object System.Drawing.Size(48,132) 
$ListBox_1.Height = 400
$ListBox_1.Width = 300
$ListBox_1.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
$ListBox_1.BackColor = "White"
$form.Controls.Add($ListBox_1)
$ListBox_1.Visible = $true

#AJOUT ELEMENTS LISTBOX1
[void]$ListBox_1.Items.Add("Informations utilisateurs")
[void]$ListBox_1.Items.Add("Vérification numéro")
[void]$ListBox_1.Items.Add("Installation d'un logiciel")
[void]$ListBox_1.Items.Add("Création d'une clé USB BOOTABLE")
[void]$ListBox_1.Items.Add("Déplacer un ordinateur (Active Directory)")

#endregion
#######################################################################################################################################################################
                                                                                # INFORMATIONS USERS #
#######################################################################################################################################################################

#region INFORMATIONsUSERS

function INFORMATIONsUSERS {

#region Structure
    
    #Fenetre
    $formADuser=New-Object System.Windows.Forms.Form
    $formADuser.Size = New-Object System.Drawing.Size(1100,1000)
    $formADuser.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $formADuser.MaximizeBox = $False
    $formADuser.MinimizeBox = $true
    $formADuser.Text = "Informations Users"
    $formADuser.StartPosition = "CenterScreen"
    $formADuser.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } })

    #ListView
    $ListViewADuser = New-Object System.Windows.Forms.ListView 
    $ListViewADuser.Location = New-Object System.Drawing.Size(600,129)
    $ListViewADuser.Size = New-Object System.Drawing.Size(435,420)
    $ListViewADuser.View = [System.Windows.Forms.View]::Details
    $ListViewADuser.HeaderStyle = "nonclickable"
    $ListViewADuser.FullRowSelect = $true
    $ListViewADuser.Columns.Add("NOM", 200)
    $ListViewADuser.Columns.Add("PRENOM", 213)
    $ListViewADuser.Enabled = $true
    $ListViewADuser.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } })
    $formADuser.Controls.Add($ListViewADuser)

    #AREAGROUPBOX INFORMATIONS
    $rechercheUserAreaGroupBox = New-Object System.Windows.Forms.GroupBox
    $rechercheUserAreaGroupBox.Location = New-Object System.Drawing.Size(50,71)
    $rechercheUserAreaGroupBox.Height = 593
    $rechercheUserAreaGroupBox.Width = 510
    $rechercheUserAreaGroupBox.Text = "Fiche utilisateur"
    $rechercheUserAreaGroupBox.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $formADuser.Controls.Add($rechercheUserAreaGroupBox)
 
    #Label NAME
    Add-Type -AssemblyName System.Windows.Forms
    $Labelname = New-Object System.Windows.Forms.Label
    $labelname.Location = New-Object System.Drawing.Size(10,20)
    $Labelname.Text = "Nom"
    $Labelname.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $Labelname.AutoSize = $false
    $Labelname.Size = New-Object System.Drawing.Size(45,20)
    $rechercheUserAreaGroupBox.Controls.Add($Labelname)
  
    #TEXTBOX NAME
    $TextboxADUuser1 = New-Object System.Windows.Forms.TextBox 
    $TextboxADUuser1.Location = New-Object System.Drawing.Size(10,40)
    $TextboxADUuser1.Size = New-Object System.Drawing.Size(190,20) 
    $TextboxADUuser1.BackColor = "white"
    $TextboxADUuser1.ReadOnly = $true
    $TextboxADUuser1.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } }) 
    $rechercheUserAreaGroupBox.Controls.Add($TextboxADUuser1)

    #Label PRENOM
    Add-Type -AssemblyName System.Windows.Forms
    $Labelprenom = New-Object System.Windows.Forms.Label
    $labelprenom.Location = New-Object System.Drawing.Size(303,20)
    $Labelprenom.Text = "Prénom"
    $Labelprenom.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $Labelprenom.AutoSize = $false
    $Labelprenom.Size = New-Object System.Drawing.Size(75,20)
    $rechercheUserAreaGroupBox.Controls.Add($Labelprenom)

    #TEXTBOX PRENOM
    $TextboxADUuser2 = New-Object System.Windows.Forms.TextBox 
    $TextboxADUuser2.Location = New-Object System.Drawing.Size(305,40)
    $TextboxADUuser2.Size = New-Object System.Drawing.Size(190,20)
    $TextboxADUuser2.BackColor = "white"
    $TextboxADUuser2.ReadOnly = $true
    $TextboxADUuser2.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } })
    $rechercheUserAreaGroupBox.Controls.Add($TextboxADUuser2)

    #Label ID
    Add-Type -AssemblyName System.Windows.Forms
    $Labelcatho = New-Object System.Windows.Forms.Label
    $Labelcatho.Location = New-Object System.Drawing.Size(10,80)
    $Labelcatho.Text = "ID"
    $Labelcatho.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $Labelcatho.AutoSize = $false
    $Labelcatho.Size = New-Object System.Drawing.Size(150,20)
    $rechercheUserAreaGroupBox.Controls.Add($Labelcatho)

    #TEXTBOX CATHO
    $TextboxADUuser3 = New-Object System.Windows.Forms.TextBox 
    $TextboxADUuser3.Location = New-Object System.Drawing.Size(10,100)
    $TextboxADUuser3.Size = New-Object System.Drawing.Size(190,20)
    $TextboxADUuser3.BackColor = "white"
    $TextboxADUuser3.ReadOnly = $true
    $TextboxADUuser3.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } })
    $rechercheUserAreaGroupBox.Controls.Add($TextboxADUuser3)

    #Label MATRICULE
    Add-Type -AssemblyName System.Windows.Forms
    $LabelMatricule = New-Object System.Windows.Forms.Label
    $LabelMatricule.Location = New-Object System.Drawing.Size(304,80)
    $LabelMatricule.Text = "Numéro Matricule "
    $LabelMatricule.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $LabelMatricule.AutoSize = $false
    $LabelMatricule.Size = New-Object System.Drawing.Size(150,20)
    $rechercheUserAreaGroupBox.Controls.Add($LabelMatricule)

    #TEXTBOX MATRICULE
    $TextboxADUuser4 = New-Object System.Windows.Forms.TextBox 
    $TextboxADUuser4.Location = New-Object System.Drawing.Size(305,100)
    $TextboxADUuser4.Size = New-Object System.Drawing.Size(190,20)
    $TextboxADUuser4.BackColor = "white"
    $TextboxADUuser4.ReadOnly = $true
    $TextboxADUuser4.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } }) 
    $rechercheUserAreaGroupBox.Controls.Add($TextboxADUuser4)

    #Label IPPHONE
    Add-Type -AssemblyName System.Windows.Forms
    $LabelIpphone = New-Object System.Windows.Forms.Label
    $LabelIpphone.Location = New-Object System.Drawing.Size(10,135)
    $LabelIpphone.Text = "Extension"
    $LabelIpphone.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $LabelIpphone.AutoSize = $false
    $LabelIpphone.Size = New-Object System.Drawing.Size(100,20)
    $rechercheUserAreaGroupBox.Controls.Add($LabelIpphone)

    #TEXTBOX IPPHONE
    $TextboxADUuser5 = New-Object System.Windows.Forms.TextBox 
    $TextboxADUuser5.Location = New-Object System.Drawing.Size(10,160)
    $TextboxADUuser5.Size = New-Object System.Drawing.Size(190,20)
    $TextboxADUuser5.BackColor = "white"
    $TextboxADUuser5.ReadOnly = $true
    $TextboxADUuser5.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } })
    $rechercheUserAreaGroupBox.Controls.Add($TextboxADUuser5)

    #Label TELEPHONE
    Add-Type -AssemblyName System.Windows.Forms
    $LabelTelephone = New-Object System.Windows.Forms.Label
    $LabelTelephone.Location = New-Object System.Drawing.Size(304,135)
    $LabelTelephone.Text = "SDA"
    $LabelTelephone.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $LabelTelephone.AutoSize = $false
    $LabelTelephone.Size = New-Object System.Drawing.Size(150,20)
    $rechercheUserAreaGroupBox.Controls.Add($LabelTelephone)

    #TEXTBOX SDA
    $TextboxADUuser6 = New-Object System.Windows.Forms.TextBox 
    $TextboxADUuser6.Location = New-Object System.Drawing.Size(305,160)
    $TextboxADUuser6.Size = New-Object System.Drawing.Size(190,20)
    $TextboxADUuser6.BackColor = "white"
    $TextboxADUuser6.ReadOnly = $true
    $TextboxADUuser6.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } })
    $rechercheUserAreaGroupBox.Controls.Add($TextboxADUuser6)

    #Label LOGIN
    Add-Type -AssemblyName System.Windows.Forms
    $LabelLOGIN = New-Object System.Windows.Forms.Label
    $LabelLOGIN.Location = New-Object System.Drawing.Size(10,200)
    $LabelLOGIN.Text = "Login"
    $LabelLOGIN.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $LabelLOGIN.AutoSize = $false
    $LabelLOGIN.Size = New-Object System.Drawing.Size(150,20)
    $rechercheUserAreaGroupBox.Controls.Add($LabelLOGIN)

    #TEXTBOX LOGIN
    $TextboxADUuser7 = New-Object System.Windows.Forms.TextBox 
    $TextboxADUuser7.Location = New-Object System.Drawing.Size(10,225)
    $TextboxADUuser7.Size = New-Object System.Drawing.Size(190,20)
    $TextboxADUuser7.BackColor = "white"
    $TextboxADUuser7.ReadOnly = $true
    $TextboxADUuser7.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } })
    $rechercheUserAreaGroupBox.Controls.Add($TextboxADUuser7)

    #Label DATE DEXPIRATION COMPTE
    Add-Type -AssemblyName System.Windows.Forms
    $LabelEXPIRACCOUNT = New-Object System.Windows.Forms.Label
    $LabelEXPIRACCOUNT.Location = New-Object System.Drawing.Size(303,200)
    $LabelEXPIRACCOUNT.Text = "Date d'expiration du compte"
    $LabelEXPIRACCOUNT.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $LabelEXPIRACCOUNT.AutoSize = $false
    $LabelEXPIRACCOUNT.Size = New-Object System.Drawing.Size(200,20)
    $rechercheUserAreaGroupBox.Controls.Add($LabelEXPIRACCOUNT)

    #TEXTBOX DATE DEXPIRATION COMPTE
    $TextboxADUuser8 = New-Object System.Windows.Forms.TextBox 
    $TextboxADUuser8.Location = New-Object System.Drawing.Size(305,225)
    $TextboxADUuser8.Size = New-Object System.Drawing.Size(190,20)
    $TextboxADUuser8.BackColor = "white"
    $TextboxADUuser8.ReadOnly = $true
    $TextboxADUuser8.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } })
    $rechercheUserAreaGroupBox.Controls.Add($TextboxADUuser8)

    #Label MAIL
    Add-Type -AssemblyName System.Windows.Forms
    $LabelMail = New-Object System.Windows.Forms.Label
    $LabelMail.Location = New-Object System.Drawing.Size(9,268)
    $LabelMail.Text = "Boite mail personnelle"
    $LabelMail.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $LabelMail.AutoSize = $false
    $LabelMail.Size = New-Object System.Drawing.Size(160,20)
    $rechercheUserAreaGroupBox.Controls.Add($LabelMail)

    #TEXTBOX MAIL
    Add-Type -AssemblyName System.Windows.Forms
    $TextboxADUuser10 = New-Object System.Windows.Forms.TextBox 
    $TextboxADUuser10.Location = New-Object System.Drawing.Size(10,295)
    $TextboxADUuser10.Size = New-Object System.Drawing.Size(240,1)
    $TextboxADUuser10.BackColor = "white"
    $TextboxADUuser10.ReadOnly = $true
    $TextboxADUuser10.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } })
    $rechercheUserAreaGroupBox.Controls.Add($TextboxADUuser10)

##########################################################################################################################

    #Label "ETAT COMPTE"
    Add-Type -AssemblyName System.Windows.Forms
    $LabelMail = New-Object System.Windows.Forms.Label
    $LabelMail.Location = New-Object System.Drawing.Size(303,268)
    $LabelMail.Text = "Etat du compte"
    $LabelMail.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $LabelMail.AutoSize = $false
    $LabelMail.Size = New-Object System.Drawing.Size(160,20)
    $rechercheUserAreaGroupBox.Controls.Add($LabelMail)

    #TEXTBOX "ETAT COMPTE"
    Add-Type -AssemblyName System.Windows.Forms
    $TextboxADUuser13 = New-Object System.Windows.Forms.TextBox 
    $TextboxADUuser13.Location = New-Object System.Drawing.Size(305,295)
    $TextboxADUuser13.Size = New-Object System.Drawing.Size(190,1)
    $TextboxADUuser13.BackColor = "white"
    $TextboxADUuser13.ReadOnly = $true
    $TextboxADUuser13.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } })
    $rechercheUserAreaGroupBox.Controls.Add($TextboxADUuser13)

##########################################################################################################################

    #Label DESCRIPTIONS
    Add-Type -AssemblyName System.Windows.Forms
    $LabelDescription = New-Object System.Windows.Forms.Label
    $LabelDescription.Location = New-Object System.Drawing.Size(220,330)
    $LabelDescription.Text = "Description"
    $LabelDescription.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $LabelDescription.AutoSize = $false
    $LabelDescription.Size = New-Object System.Drawing.Size(150,20)
    $rechercheUserAreaGroupBox.Controls.Add($LabelDescription)

    #TEXTBOX DESCRIPTIONS
    Add-Type -AssemblyName System.Windows.Forms
    $TextboxADUuser9 = New-Object System.Windows.Forms.TextBox 
    $TextboxADUuser9.Location = New-Object System.Drawing.Size(10,355)
    $TextboxADUuser9.Size = New-Object System.Drawing.Size(485,1)
    $TextboxADUuser9.BackColor = "white"
    $TextboxADUuser9.ReadOnly = $true
    $TextboxADUuser9.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } })
    $rechercheUserAreaGroupBox.Controls.Add($TextboxADUuser9)

    #LABEL BARRE DE RECHERCHE
    Add-Type -AssemblyName System.Windows.Forms
    $LabelBarreDeRecherche = New-Object System.Windows.Forms.Label
    $LabelBarreDeRecherche.Location = New-Object System.Drawing.Size(600,75)
    $LabelBarreDeRecherche.Text = "Rechercher un utilisateur (nom ou prénom) :"
    $LabelBarreDeRecherche.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $LabelBarreDeRecherche.AutoSize = $false
    $LabelBarreDeRecherche.Size = New-Object System.Drawing.Size(310,18)
    $formADuser.Controls.Add($LabelBarreDeRecherche)

    #TEXTBOX BARRE DE RECHERCHE
    $TextboxADUuser12 = New-Object System.Windows.Forms.TextBox 
    $TextboxADUuser12.Location = New-Object System.Drawing.Size(600,94)
    $TextboxADUuser12.Size = New-Object System.Drawing.Size(340,25)
    $TextboxADUuser12.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $TextboxADUuser12.Enabled = $true
    $TextboxADUuser12.Multiline = $false
    $TextboxADUuser12.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } })
    $TextboxADUuser12.Add_KeyDown({ if($_.KeyCode -eq "Enter") { $ListViewADuser.Items.Clear() } })
    $TextboxADUuser12.Add_KeyDown({ if($_.KeyCode -eq "Enter") { BARREDERECHERCHE } })
    $formADuser.Controls.Add($TextboxADUuser12)

    #BUTTONS BARRE DE RECHERCHE
    Add-Type -AssemblyName System.Windows.Forms
    $buttonBarreDeRecherche = New-Object System.Windows.Forms.Button
    $buttonBarreDeRecherche.Text = "Go"
    $buttonBarreDeRecherche.Font = New-Object System.Drawing.Font("Arial",9,[System.Drawing.FontStyle]::Bold)
    $buttonBarreDeRecherche.Width = 80
    $buttonBarreDeRecherche.Height = 30
    $buttonBarreDeRecherche.Location = New-Object System.Drawing.Size(955,92)
    $buttonBarreDeRecherche.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } })
    $buttonBarreDeRecherche.add_Click({$ListViewADuser.Items.Clear()})
    $buttonBarreDeRecherche.add_Click({$TextboxADUuser1.Clear()})
    $buttonBarreDeRecherche.add_Click({$TextboxADUuser2.Clear()})
    $buttonBarreDeRecherche.add_Click({$TextboxADUuser3.Clear()})
    $buttonBarreDeRecherche.add_Click({$TextboxADUuser4.Clear()})
    $buttonBarreDeRecherche.add_Click({$TextboxADUuser5.Clear()})
    $buttonBarreDeRecherche.add_Click({$TextboxADUuser6.Clear()})
    $buttonBarreDeRecherche.add_Click({$TextboxADUuser7.Clear()})
    $buttonBarreDeRecherche.add_Click({$TextboxADUuser8.Clear()})
    $buttonBarreDeRecherche.add_Click({$TextboxADUuser9.Clear()})
    $buttonBarreDeRecherche.add_Click({$TextboxADUuser10.Clear()})
    $buttonBarreDeRecherche.add_Click({$TextboxADUuser11.Clear()})
    $buttonBarreDeRecherche.add_Click({$TextboxModifMatricule1.Clear()})
    $buttonBarreDeRecherche.add_Click({$TextboxModifMatricule2.Clear()})
    $buttonBarreDeRecherche.add_Click({$TextboxModifEXTENSION1.Clear()})
    $buttonBarreDeRecherche.add_Click({$TextboxModifEXTENSION2.Clear()})
    $buttonBarreDeRecherche.add_Click({$TextboxModifSDA1.Clear()})
    $buttonBarreDeRecherche.add_Click({$TextboxModifSDA2.Clear()})
    $buttonBarreDeRecherche.add_Click({$TextboxADUuser13.Clear()})
    $buttonBarreDeRecherche.add_Click({$ModifDateExpiration1.Clear()})
    $buttonBarreDeRecherche.add_Click({$ModifDateExpiration2.Clear()})
    $buttonBarreDeRecherche.add_Click({BARREDERECHERCHE})
    $formADuser.Controls.Add($buttonBarreDeRecherche)

    #LABEL GROUPES ET BOITES MAILS
    Add-Type -AssemblyName System.Windows.Forms
    $LabelMembre = New-Object System.Windows.Forms.Label
    $LabelMembre.Location = New-Object System.Drawing.Size(185,390)
    $LabelMembre.Text = "Groupes et Boites mails"
    $LabelMembre.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $LabelMembre.AutoSize = $false
    $LabelMembre.Size = New-Object System.Drawing.Size(170,20)
    $rechercheUserAreaGroupBox.Controls.Add($LabelMembre)
          
    #TEXTBOX GROUPES ET BOITES MAILS
    Add-Type -AssemblyName System.Windows.Forms
    $TextboxADUuser11 = New-Object System.Windows.Forms.TextBox 
    $TextboxADUuser11.Location = New-Object System.Drawing.Size(10,415)
    $TextboxADUuser11.Size = New-Object System.Drawing.Size(485,165)
    $TextboxADUuser11.BackColor = "white"
    $TextboxADUuser11.ReadOnly = $true
    $TextboxADUuser11.ScrollBars = "Vertical"
    $TextboxADUuser11.Multiline = $true
    $TextboxADUuser11.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } })
    $rechercheUserAreaGroupBox.Controls.Add($TextboxADUuser11)

#endregion Structure

####################################################################################################################################################################### 
                                                                                     #GROUPES#
#######################################################################################################################################################################

#region Groupes 2, 3 et 4

    #GROUPE 2 BUTTONS
    $Group2 = New-Object System.Windows.Forms.GroupBox
    $Group2.Location = New-Object System.Drawing.Size(594,550)
    $Group2.Height = 115
    $Group2.Width = 450
    $Group2.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)

    #GROUPE 3 ACTIONS USERS
    $Group3 = New-Object System.Windows.Forms.GroupBox
    $Group3.Location = New-Object System.Drawing.Size(50,680)
    $Group3.Height = 265
    $Group3.Width = 850
    $Group3.Text = "Modifications fiche utilisateur"
    $Group3.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)

    #GROUPE 4 ACTIONS COMPTE ACTIVATION|DESACTIVATION
    $Group4 = New-Object System.Windows.Forms.GroupBox
    $Group4.Location = New-Object System.Drawing.Size(485,780)
    $Group4.Height = 110
    $Group4.Width = 402
  
  #############################################    

    #LABEL MODIF MATRICULE
    Add-Type -AssemblyName System.Windows.Forms
    $LabelModifMatricule = New-Object System.Windows.Forms.Label
    $LabelModifMatricule.Location = New-Object System.Drawing.Size(8,35)
    $LabelModifMatricule.Text = "Matricule"
    $LabelModifMatricule.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $LabelModifMatricule.AutoSize = $false
    $LabelModifMatricule.Size = New-Object System.Drawing.Size(170,20)
    $Group3.Controls.Add($LabelModifMatricule)
    
    #TEXTBOX MODIF MATRICULE (LOGIN)
    Add-Type -AssemblyName System.Windows.Forms
    $TextboxModifMatricule1 = New-Object System.Windows.Forms.TextBox 
    $TextboxModifMatricule1.Location = New-Object System.Drawing.Size(10,55)
    $TextboxModifMatricule1.Size = New-Object System.Drawing.Size(170,20)
    $TextboxModifMatricule1.BackColor = "white"
    $TextboxModifMatricule1.ReadOnly = $true
    $TextboxModifMatricule1.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } })
    $Group3.Controls.Add($TextboxModifMatricule1)
 
    #TEXTBOX MODIF MATRICULE (MATRICULE)
    Add-Type -AssemblyName System.Windows.Forms
    $TextboxModifMatricule2 = New-Object System.Windows.Forms.TextBox 
    $TextboxModifMatricule2.Location = New-Object System.Drawing.Size(195,55)
    $TextboxModifMatricule2.MaxLength = 5
    $TextboxModifMatricule2.Size = New-Object System.Drawing.Size(90,20)
    $TextboxModifMatricule2.BackColor = "white"
    $TextboxModifMatricule2.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } })
    $Group3.Controls.Add($TextboxModifMatricule2)
   
    #BUTTONS ACTION MATRICULE
    Add-Type -AssemblyName System.Windows.Forms
    $buttonModifMatricule= New-Object System.Windows.Forms.Button
    $buttonModifMatricule.Text = "Executer"
    $buttonModifMatricule.Font = New-Object System.Drawing.Font("Arial",9,[System.Drawing.FontStyle]::Regular)
    $buttonModifMatricule.Width = 110
    $buttonModifMatricule.Height = 30
    $buttonModifMatricule.Location = New-Object System.Drawing.Size(300,52)
    $buttonModifMatricule.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } })
    $buttonModifMatricule.add_Click({MATRICULE})
    $Group3.Controls.Add($buttonModifMatricule)

    #LABEL MODIF SDA PHONE
    Add-Type -AssemblyName System.Windows.Forms
    $LabelModifSDA = New-Object System.Windows.Forms.Label
    $LabelModifSDA.Location = New-Object System.Drawing.Size(8,95)
    $LabelModifSDA.Text = "Modifier SDA"
    $LabelModifSDA.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $LabelModifSDA.AutoSize = $false
    $LabelModifSDA.Size = New-Object System.Drawing.Size(170,20)
    $Group3.Controls.Add($LabelModifSDA)
    
    #TEXTBOX MODIF SDA PHONE (LOGIN)
    Add-Type -AssemblyName System.Windows.Forms
    $TextboxModifSDA1 = New-Object System.Windows.Forms.TextBox 
    $TextboxModifSDA1.Location = New-Object System.Drawing.Size(10,115)
    $TextboxModifSDA1.Size = New-Object System.Drawing.Size(170,20)
    $TextboxModifSDA1.BackColor = "white"
    $TextboxModifSDA1.ReadOnly = $true
    $TextboxModifSDA1.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } })
    $Group3.Controls.Add($TextboxModifSDA1)
    
    #TEXTBOX MODIF SDA (SDA)
    Add-Type -AssemblyName System.Windows.Forms
    $TextboxModifSDA2 = New-Object System.Windows.Forms.TextBox 
    $TextboxModifSDA2.Location = New-Object System.Drawing.Size(195,115)
    $TextboxModifSDA2.MaxLength = 10
    $TextboxModifSDA2.Size = New-Object System.Drawing.Size(90,20)
    $TextboxModifSDA2.BackColor = "white"
    $TextboxModifSDA2.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } })
    $Group3.Controls.Add($TextboxModifSDA2)

    #BUTTONS ACTION SDA
    Add-Type -AssemblyName System.Windows.Forms
    $buttonModifSDA= New-Object System.Windows.Forms.Button
    $buttonModifSDA.Text = "Executer"
    $buttonModifSDA.Font = New-Object System.Drawing.Font("Arial",9,[System.Drawing.FontStyle]::Regular)
    $buttonModifSDA.Width = 110
    $buttonModifSDA.Height = 30
    $buttonModifSDA.Location = New-Object System.Drawing.Size(300,112)
    $buttonModifSDA.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } })
    $buttonModifSDA.add_Click({MODIFSDA})
    $Group3.Controls.Add($buttonModifSDA)


    #LABEL MODIF EXTENSION
    Add-Type -AssemblyName System.Windows.Forms
    $LabelModifExtension = New-Object System.Windows.Forms.Label
    $LabelModifExtension.Location = New-Object System.Drawing.Size(8,155)
    $LabelModifExtension.Text = "Modifier EXTENSION"
    $LabelModifExtension.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $LabelModifExtension.AutoSize = $false
    $LabelModifExtension.Size = New-Object System.Drawing.Size(170,20)
    $Group3.Controls.Add($LabelModifExtension)
    
  
    #TEXTBOX MODIF EXTENSION (LOGIN)
    Add-Type -AssemblyName System.Windows.Forms
    $TextboxModifEXTENSION1 = New-Object System.Windows.Forms.TextBox 
    $TextboxModifEXTENSION1.Location = New-Object System.Drawing.Size(10,175)
    $TextboxModifEXTENSION1.Size = New-Object System.Drawing.Size(170,20)
    $TextboxModifEXTENSION1.BackColor = "white"
    $TextboxModifEXTENSION1.ReadOnly = $true
    $TextboxModifEXTENSION1.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } })
    $Group3.Controls.Add($TextboxModifEXTENSION1)

    
    #TEXTBOX MODIF EXTENSION (EXTENSION)
    Add-Type -AssemblyName System.Windows.Forms
    $TextboxModifEXTENSION2 = New-Object System.Windows.Forms.TextBox 
    $TextboxModifEXTENSION2.Location = New-Object System.Drawing.Size(195,175)
    $TextboxModifEXTENSION2.MaxLength = 4
    $TextboxModifEXTENSION2.Size = New-Object System.Drawing.Size(90,20)
    $TextboxModifEXTENSION2.BackColor = "white"
    $TextboxModifEXTENSION2.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } })
    $Group3.Controls.Add($TextboxModifEXTENSION2)  

    #LABEL MODIF DATE EXPIRATION
    Add-Type -AssemblyName System.Windows.Forms
    $ModifDateExpiration = New-Object System.Windows.Forms.Label
    $ModifDateExpiration.Location = New-Object System.Drawing.Size(433,35)
    $ModifDateExpiration.Text = "Modifier date d'expiration"
    $ModifDateExpiration.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
    $ModifDateExpiration.AutoSize = $false
    $ModifDateExpiration.Size = New-Object System.Drawing.Size(174,20)
    $Group3.Controls.Add($ModifDateExpiration)
     
    #TEXTBOX MODIF DATE EXPIRATION (LOGIN)
    Add-Type -AssemblyName System.Windows.Forms
    $ModifDateExpiration1 = New-Object System.Windows.Forms.TextBox 
    $ModifDateExpiration1.Location = New-Object System.Drawing.Size(435,55)
    $ModifDateExpiration1.Size = New-Object System.Drawing.Size(170,20)
    $ModifDateExpiration1.BackColor = "white"
    $ModifDateExpiration1.ReadOnly = $true
    $ModifDateExpiration1.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } })
    $Group3.Controls.Add($ModifDateExpiration1)

    
    #TEXTBOX MODIF DATE EXPIRATION (DATE)
    Add-Type -AssemblyName System.Windows.Forms
    $ModifDateExpiration2 = New-Object System.Windows.Forms.TextBox 
    $ModifDateExpiration2.Location = New-Object System.Drawing.Size(624,55)
    $ModifDateExpiration2.MaxLength = 8
    $ModifDateExpiration2.Size = New-Object System.Drawing.Size(90,20)
    $ModifDateExpiration2.BackColor = "white"
    $ModifDateExpiration2.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } })
    $Group3.Controls.Add($ModifDateExpiration2)

   
    #BUTTONS ACTION MODIF DATE EXPIRATION
    Add-Type -AssemblyName System.Windows.Forms
    $ModifDateExpiration3= New-Object System.Windows.Forms.Button
    $ModifDateExpiration3.Text = "Executer"
    $ModifDateExpiration3.Font = New-Object System.Drawing.Font("Arial",9,[System.Drawing.FontStyle]::Regular)
    $ModifDateExpiration3.Width = 110
    $ModifDateExpiration3.Height = 30
    $ModifDateExpiration3.Location = New-Object System.Drawing.Size(728,52)
    $ModifDateExpiration3.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } })
    $ModifDateExpiration3.add_Click({COMPTEEXPIRATION})
    $Group3.Controls.Add($ModifDateExpiration3)


    #BUTTONS ACTION MODIF ACTIVATION COMPTE
    Add-Type -AssemblyName System.Windows.Forms
    $ModifActivationCompte= New-Object System.Windows.Forms.Button
    $ModifActivationCompte.Text = "Activer compte"
    $ModifActivationCompte.Font = New-Object System.Drawing.Font("Arial",9,[System.Drawing.FontStyle]::Regular)
    $ModifActivationCompte.Width = 190
    $ModifActivationCompte.Height = 40
    $ModifActivationCompte.Location = New-Object System.Drawing.Size(494,795)
    $ModifActivationCompte.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } })
    $ModifActivationCompte.add_Click({ACTIVATIONCOMPTE})
    $formADuser.Controls.Add($ModifActivationCompte)
    
    #BUTTONS ACTION MODIF DESACTIVATION COMPTE
    Add-Type -AssemblyName System.Windows.Forms
    $ModifDesactivationCompte= New-Object System.Windows.Forms.Button
    $ModifDesactivationCompte.Text = "Désactiver compte"
    $ModifDesactivationCompte.Font = New-Object System.Drawing.Font("Arial",9,[System.Drawing.FontStyle]::Regular)
    $ModifDesactivationCompte.Width = 190
    $ModifDesactivationCompte.Height = 40
    $ModifDesactivationCompte.Location = New-Object System.Drawing.Size(494,842)
    $ModifDesactivationCompte.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } })
    $ModifDesactivationCompte.add_Click({DESACTIVATIONCOMPTE})
    $formADuser.Controls.Add($ModifDesactivationCompte)
   
    #BUTTONS ACTION AJOUTER BOITE MAIL
    Add-Type -AssemblyName System.Windows.Forms
    $ModifAjouterMail= New-Object System.Windows.Forms.Button
    $ModifAjouterMail.Text = "Ajouter Boite mails"
    $ModifAjouterMail.Font = New-Object System.Drawing.Font("Arial",9,[System.Drawing.FontStyle]::Regular)
    $ModifAjouterMail.Width = 190
    $ModifAjouterMail.Height = 40
    $ModifAjouterMail.Location = New-Object System.Drawing.Size(690,795)
    $ModifAjouterMail.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } })
    $ModifAjouterMail.add_Click({AJOUTBAL})
    $formADuser.Controls.Add($ModifAjouterMail)


    #BUTTONS ACTION AJOUTER VPN
    Add-Type -AssemblyName System.Windows.Forms
    $ModifAjouterVPN= New-Object System.Windows.Forms.Button
    $ModifAjouterVPN.Text = "Ajouter VPN"
    $ModifAjouterVPN.Font = New-Object System.Drawing.Font("Arial",9,[System.Drawing.FontStyle]::Regular)
    $ModifAjouterVPN.Width = 190
    $ModifAjouterVPN.Height = 40
    $ModifAjouterVPN.Location = New-Object System.Drawing.Size(690,842)
    $ModifAjouterVPN.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } })
    $ModifAjouterVPN.add_Click({AJOUTVPN})
    $formADuser.Controls.Add($ModifAjouterVPN)

#############################################    
    #AFFICHAGE GROUPE 4
    $formADuser.Controls.Add($Group4)

    #AFFICHAGE GROUPE 3
    $formADuser.Controls.Add($Group3)
#############################################    

#endregion

####################################################################################################################################################################### 
                                                                                # BUTTONS + EVENEMENTS #
####################################################################################################################################################################### 

#region Buttons


    #BUTTONS ADMINISTRATIF
    Add-Type -AssemblyName System.Windows.Forms
    $buttonAdmin = New-Object System.Windows.Forms.Button   
    $buttonAdmin.Text = "Administratif"
    $buttonAdmin.Font = New-Object System.Drawing.Font("Arial",9,[System.Drawing.FontStyle]::Regular)
    $buttonAdmin.Width = 130
    $buttonAdmin.Height = 40
    $buttonAdmin.Location = New-Object System.Drawing.Size(600,565)
    $buttonAdmin.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } })
    $formADuser.Controls.Add($buttonAdmin)
    $buttonAdmin.add_Click({$ListViewADuser.Items.Clear()})
    $buttonAdmin.add_Click({$TextboxADUuser1.Clear()})
    $buttonAdmin.add_Click({$TextboxADUuser2.Clear()})
    $buttonAdmin.add_Click({$TextboxADUuser3.Clear()})
    $buttonAdmin.add_Click({$TextboxADUuser4.Clear()})
    $buttonAdmin.add_Click({$TextboxADUuser5.Clear()})
    $buttonAdmin.add_Click({$TextboxADUuser6.Clear()})
    $buttonAdmin.add_Click({$TextboxADUuser7.Clear()})
    $buttonAdmin.add_Click({$TextboxADUuser8.Clear()})
    $buttonAdmin.add_Click({$TextboxADUuser9.Clear()})
    $buttonAdmin.add_Click({$TextboxADUuser10.Clear()})
    $buttonAdmin.add_Click({$TextboxADUuser11.Clear()})
    $buttonAdmin.add_Click({$TextboxModifMatricule1.Clear()})
    $buttonAdmin.add_Click({$TextboxModifMatricule2.Clear()})
    $buttonAdmin.add_Click({$TextboxModifEXTENSION1.Clear()})
    $buttonAdmin.add_Click({$TextboxModifEXTENSION2.Clear()})
    $buttonAdmin.add_Click({$TextboxModifSDA1.Clear()})
    $buttonAdmin.add_Click({$TextboxModifSDA2.Clear()})
    $buttonAdmin.add_Click({$TextboxADUuser13.Clear()})
    $buttonAdmin.add_Click({$ModifDateExpiration1.Clear()})
    $buttonAdmin.add_Click({$ModifDateExpiration2.Clear()}) 
    $buttonAdmin.add_Click({Administratif})


    #BUTTONS INSTITUTS
    Add-Type -AssemblyName System.Windows.Forms
    $buttonInstituts= New-Object System.Windows.Forms.Button
    $buttonInstituts.Text = "Entreprise"
    $buttonInstituts.Font = New-Object System.Drawing.Font("Arial",9,[System.Drawing.FontStyle]::Regular)
    $buttonInstituts.Width = 130
    $buttonInstituts.Height = 40
    $buttonInstituts.Location = New-Object System.Drawing.Size(756,565)
    $buttonInstituts.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } })
    $buttonInstituts.add_Click({$ListViewADuser.Items.Clear()})
    $buttonInstituts.add_Click({$TextboxADUuser1.Clear()})
    $buttonInstituts.add_Click({$TextboxADUuser2.Clear()})
    $buttonInstituts.add_Click({$TextboxADUuser3.Clear()})
    $buttonInstituts.add_Click({$TextboxADUuser4.Clear()})
    $buttonInstituts.add_Click({$TextboxADUuser5.Clear()})
    $buttonInstituts.add_Click({$TextboxADUuser6.Clear()})
    $buttonInstituts.add_Click({$TextboxADUuser7.Clear()})
    $buttonInstituts.add_Click({$TextboxADUuser8.Clear()})
    $buttonInstituts.add_Click({$TextboxADUuser9.Clear()})
    $buttonInstituts.add_Click({$TextboxADUuser10.Clear()})
    $buttonInstituts.add_Click({$TextboxADUuser11.Clear()})
    $buttonInstituts.add_Click({$TextboxModifMatricule1.Clear()})
    $buttonInstituts.add_Click({$TextboxModifMatricule2.Clear()})
    $buttonInstituts.add_Click({$TextboxModifEXTENSION1.Clear()})
    $buttonInstituts.add_Click({$TextboxModifEXTENSION2.Clear()})
    $buttonInstituts.add_Click({$TextboxModifSDA1.Clear()})
    $buttonInstituts.add_Click({$TextboxModifSDA2.Clear()})
    $buttonInstituts.add_Click({$TextboxADUuser13.Clear()})
    $buttonInstituts.add_Click({$ModifDateExpiration1.Clear()})
    $buttonInstituts.add_Click({$ModifDateExpiration2.Clear()})
    $buttonInstituts.add_Click({Personnel})
    $formADuser.Controls.Add($buttonInstituts)

    #BUTTONS SERVICE INFORMATIQUE
    Add-Type -AssemblyName System.Windows.Forms
    $buttonSI= New-Object System.Windows.Forms.Button
    $buttonSI.Text = "Service Informatique"
    $buttonSI.Font = New-Object System.Drawing.Font("Arial",9,[System.Drawing.FontStyle]::Regular)
    $buttonSI.Width = 130
    $buttonSI.Height = 40
    $buttonSI.Location = New-Object System.Drawing.Size(905,565)
    $buttonSI.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } })
    $buttonSI.add_Click({$ListViewADuser.Items.Clear()})
    $buttonSI.add_Click({$TextboxADUuser1.Clear()})
    $buttonSI.add_Click({$TextboxADUuser2.Clear()})
    $buttonSI.add_Click({$TextboxADUuser3.Clear()})
    $buttonSI.add_Click({$TextboxADUuser4.Clear()})
    $buttonSI.add_Click({$TextboxADUuser5.Clear()})
    $buttonSI.add_Click({$TextboxADUuser6.Clear()})
    $buttonSI.add_Click({$TextboxADUuser7.Clear()})
    $buttonSI.add_Click({$TextboxADUuser8.Clear()})
    $buttonSI.add_Click({$TextboxADUuser9.Clear()})
    $buttonSI.add_Click({$TextboxADUuser10.Clear()})
    $buttonSI.add_Click({$TextboxADUuser11.Clear()})
    $buttonSI.add_Click({$TextboxModifMatricule1.Clear()})
    $buttonSI.add_Click({$TextboxModifMatricule2.Clear()})
    $buttonSI.add_Click({$TextboxModifEXTENSION1.Clear()})
    $buttonSI.add_Click({$TextboxModifEXTENSION2.Clear()})
    $buttonSI.add_Click({$TextboxModifSDA1.Clear()})
    $buttonSI.add_Click({$TextboxModifSDA2.Clear()})
    $buttonSI.add_Click({$TextboxADUuser13.Clear()})
    $buttonSI.add_Click({$ModifDateExpiration1.Clear()})
    $buttonSI.add_Click({$ModifDateExpiration2.Clear()})
    $buttonSI.add_Click({SERVICEINFO})
    $formADuser.Controls.Add($buttonSI)

    #BUTTONS ADMIN COMPTES DESACTIVES
    Add-Type -AssemblyName System.Windows.Forms
    $buttonAdminlocked = New-Object System.Windows.Forms.Button
    $buttonAdminlocked.Text = "Comptes Admin Désactivés"
    $buttonAdminlocked.Font = New-Object System.Drawing.Font("Arial",9,[System.Drawing.FontStyle]::Regular)
    $buttonAdminlocked.Width = 130
    $buttonAdminlocked.Height = 40
    $buttonAdminlocked.Location = New-Object System.Drawing.Size(600,615)
    $buttonAdminlocked.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } })
    $buttonAdminlocked.add_Click({$ListViewADuser.Items.Clear()})
    $buttonAdminlocked.add_Click({$TextboxADUuser1.Clear()})
    $buttonAdminlocked.add_Click({$TextboxADUuser2.Clear()})
    $buttonAdminlocked.add_Click({$TextboxADUuser3.Clear()})
    $buttonAdminlocked.add_Click({$TextboxADUuser4.Clear()})
    $buttonAdminlocked.add_Click({$TextboxADUuser5.Clear()})
    $buttonAdminlocked.add_Click({$TextboxADUuser6.Clear()})
    $buttonAdminlocked.add_Click({$TextboxADUuser7.Clear()})
    $buttonAdminlocked.add_Click({$TextboxADUuser8.Clear()})
    $buttonAdminlocked.add_Click({$TextboxADUuser9.Clear()})
    $buttonAdminlocked.add_Click({$TextboxADUuser10.Clear()})
    $buttonAdminlocked.add_Click({$TextboxADUuser11.Clear()})
    $buttonAdminlocked.add_Click({$TextboxModifMatricule1.Clear()})
    $buttonAdminlocked.add_Click({$TextboxModifMatricule2.Clear()})
    $buttonAdminlocked.add_Click({$TextboxModifEXTENSION1.Clear()})
    $buttonAdminlocked.add_Click({$TextboxModifEXTENSION2.Clear()})
    $buttonAdminlocked.add_Click({$TextboxModifSDA1.Clear()})
    $buttonAdminlocked.add_Click({$TextboxModifSDA2.Clear()})
    $buttonAdminlocked.add_Click({$TextboxADUuser13.Clear()})
    $buttonAdminlocked.add_Click({$ModifDateExpiration1.Clear()})
    $buttonAdminlocked.add_Click({$ModifDateExpiration2.Clear()})
    $buttonAdminlocked.add_Click({LOCKEDADMIN_PROF_PARTEN})
    $formADuser.Controls.Add($buttonAdminlocked) 

    #BUTTONS ETUDIANTS
    Add-Type -AssemblyName System.Windows.Forms
    $buttonETUDIANTS= New-Object System.Windows.Forms.Button
    $buttonETUDIANTS.Text = "Utilisateurs"
    $buttonETUDIANTS.Font = New-Object System.Drawing.Font("Arial",9,[System.Drawing.FontStyle]::Regular)
    $buttonETUDIANTS.Width = 130
    $buttonETUDIANTS.Height = 40
    $buttonETUDIANTS.Location = New-Object System.Drawing.Size(756,615)
    $buttonETUDIANTS.Add_KeyDown({ if($_.KeyCode -eq "Escape") {$formADuser.Close()}})
    $buttonETUDIANTS.add_Click({Etudiants})   
    $formADuser.Controls.Add($buttonETUDIANTS)


    #BUTTONS CLEAR
    Add-Type -AssemblyName System.Windows.Forms
    $buttonCLEAR= New-Object System.Windows.Forms.Button
    $buttonCLEAR.Text = "CLEAR"
    $buttonCLEAR.Font = New-Object System.Drawing.Font("Arial",9,[System.Drawing.FontStyle]::Regular)
    $buttonCLEAR.Width = 130
    $buttonCLEAR.Height = 40
    $buttonCLEAR.Location = New-Object System.Drawing.Size(905,615)
    $buttonCLEAR.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } })
    $buttonCLEAR.add_Click({$ListViewADuser.Items.Clear()})
    $buttonCLEAR.add_Click({$TextboxADUuser1.Clear()})
    $buttonCLEAR.add_Click({$TextboxADUuser2.Clear()})
    $buttonCLEAR.add_Click({$TextboxADUuser3.Clear()})
    $buttonCLEAR.add_Click({$TextboxADUuser4.Clear()})
    $buttonCLEAR.add_Click({$TextboxADUuser5.Clear()})
    $buttonCLEAR.add_Click({$TextboxADUuser6.Clear()})
    $buttonCLEAR.add_Click({$TextboxADUuser7.Clear()})
    $buttonCLEAR.add_Click({$TextboxADUuser8.Clear()})
    $buttonCLEAR.add_Click({$TextboxADUuser9.Clear()})
    $buttonCLEAR.add_Click({$TextboxADUuser10.Clear()})
    $buttonCLEAR.add_Click({$TextboxADUuser11.Clear()})
    $buttonCLEAR.add_Click({$TextboxADUuser12.Clear()})
    $buttonCLEAR.add_Click({$TextboxModifMatricule1.Clear()})
    $buttonCLEAR.add_Click({$TextboxModifMatricule2.Clear()})
    $buttonCLEAR.add_Click({$TextboxModifEXTENSION1.Clear()})
    $buttonCLEAR.add_Click({$TextboxModifEXTENSION2.Clear()})
    $buttonCLEAR.add_Click({$TextboxModifSDA1.Clear()})
    $buttonCLEAR.add_Click({$TextboxModifSDA2.Clear()})
    $buttonCLEAR.add_Click({$TextboxADUuser13.Clear()})
    $buttonCLEAR.add_Click({$ModifDateExpiration1.Clear()})
    $buttonCLEAR.add_Click({$ModifDateExpiration2.Clear()})
    $formADuser.Controls.Add($buttonCLEAR)
   
    #BUTTONS INFORMATIONS TOTALES
    Add-Type -AssemblyName System.Windows.Forms
    $buttonInfosTOTALES= New-Object System.Windows.Forms.Button
    $buttonInfosTOTALES.Text = "INFOS TOTALES"
    $buttonInfosTOTALES.Font = New-Object System.Drawing.Font("Arial",9,[System.Drawing.FontStyle]::Regular)
    $buttonInfosTOTALES.Width = 130
    $buttonInfosTOTALES.Height = 40
    $buttonInfosTOTALES.Location = New-Object System.Drawing.Size(905,715)
    $buttonInfosTOTALES.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } })
    $buttonInfosTOTALES.add_Click({})
    $formADuser.Controls.Add($buttonInfosTOTALES)

    #BUTTONS FERMER FENETRE INFORMATIONS USERS
    Add-Type -AssemblyName System.Windows.Forms
    $buttonCloseWindowUsers= New-Object System.Windows.Forms.Button
    $buttonCloseWindowUsers.Text = "Fermer Fenêtre"
    $buttonCloseWindowUsers.Font = New-Object System.Drawing.Font("Arial",9,[System.Drawing.FontStyle]::Regular)
    $buttonCloseWindowUsers.Width = 130
    $buttonCloseWindowUsers.Height = 40
    $buttonCloseWindowUsers.Location = New-Object System.Drawing.Size(905,795)
    $buttonCloseWindowUsers.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } })
    $buttonCloseWindowUsers.add_Click({$formADuser.Close()})
    $formADuser.Controls.Add($buttonCloseWindowUsers)

    #BUTTONS FERMER FENETRES INFORMATIONS USERS + MENU GENERAL
    Add-Type -AssemblyName System.Windows.Forms
    $buttonCloseWindowUsersGENERAL= New-Object System.Windows.Forms.Button
    $buttonCloseWindowUsersGENERAL.Text = "Fermer PROGRAMME"
    $buttonCloseWindowUsersGENERAL.Font = New-Object System.Drawing.Font("Arial",9,[System.Drawing.FontStyle]::Regular)
    $buttonCloseWindowUsersGENERAL.Width = 130
    $buttonCloseWindowUsersGENERAL.Height = 40
    $buttonCloseWindowUsersGENERAL.Location = New-Object System.Drawing.Size(905,875)
    $buttonCloseWindowUsersGENERAL.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } })
    $buttonCloseWindowUsersGENERAL.add_Click({$formADuser.Close()})
    $buttonCloseWindowUsersGENERAL.add_Click({$form.Close()})
    $formADuser.Controls.Add($buttonCloseWindowUsersGENERAL)

    #BUTTONS ACTION EXTENSION
    Add-Type -AssemblyName System.Windows.Forms
    $buttonModifEXTENSION= New-Object System.Windows.Forms.Button
    $buttonModifEXTENSION.Text = "Executer"
    $buttonModifEXTENSION.Font = New-Object System.Drawing.Font("Arial",9,[System.Drawing.FontStyle]::Regular)
    $buttonModifEXTENSION.Width = 110
    $buttonModifEXTENSION.Height = 30
    $buttonModifEXTENSION.Location = New-Object System.Drawing.Size(300,172)
    $buttonModifEXTENSION.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formADuser.Close() } })
    $buttonModifEXTENSION.add_Click({MODIFEXTENSION})
    $Group3.Controls.Add($buttonModifEXTENSION)

##############################################
    #AFFICHAGE GROUPE 2 | BUTTONS ACTION
    $formADuser.Controls.Add($Group2)
##############################################
#endregion Buttons

####################################################################################################################################################################### 
                                                                            # FUNCTIONS #
####################################################################################################################################################################### 

#region Functions

Function Administratif{
       
        $usersad = Get-ADUser -Filter * -SearchBase $UO -Server $Server -Properties * | Sort-Object 
         
        $usersad | ForEach {
                  
                  If($_.Surname -eq $null) {$ListViewADuserITEM = New-Object System.Windows.Forms.ListViewItem("")}
                  else {$ListViewADuserITEM = New-Object System.Windows.Forms.ListViewItem($_.Surname)}
                                 
                  If($_.GivenName -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.GivenName)}
                               
                  If($_.uidnumber -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.uidnumber)}
    
                  If($_.EmployeeID -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.EmployeeID)}
                                    
                  If($_.ipphone -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.ipphone)}
             
                  If($_.TelephoneNumber -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.TelephoneNumber)}   
                                                 
                  If($_.SamAccountName -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.SamAccountName)}               
                
                  If($_.AccountExpires -eq 0){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add("$(Get-Date([datetime]::FromFileTime($_.AccountExpires)) -Format "d MMM yyyy")")}
                
                  If($_.Description -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.Description)}  
                                    
                  If($_.Description -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.mail)}  
              
                  $ListViewADuser.Items.Add($ListViewADuserITEM)     

                   }                                  
}


Function Personnel {
        
        $usersad = Get-ADUser -Filter * -SearchBase $UO -Server $Domaine -Properties * | Sort-Object 
         
        $usersad | ForEach {
                  
                  If($_.Surname -eq $null) {$ListViewADuserITEM = New-Object System.Windows.Forms.ListViewItem("")}
                  else {$ListViewADuserITEM = New-Object System.Windows.Forms.ListViewItem($_.Surname)}
                                    
                  If($_.GivenName -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.GivenName)}
                                   
                  If($_.uidnumber -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.uidnumber)}
                  
                  If($_.EmployeeID -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.EmployeeID)}                 
                                  
                  If($_.ipphone -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.ipphone)}
                 
                  If($_.TelephoneNumber -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.TelephoneNumber)}   
                                                
                  If($_.SamAccountName -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.SamAccountName)}               
                
                  If($_.AccountExpires -eq 0){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add("$(Get-Date([datetime]::FromFileTime($_.AccountExpires)) -Format "d MMM yyyy")")}
                 
                  If($_.Description -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.Description)}  
                                     
                  If($_.Description -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.mail)}  

                  $ListViewADuser.Items.Add($ListViewADuserITEM)
       

                  }
}


Function SERVICEINFO {
        
        $usersad = Get-ADUser -Filter * -SearchBase $UO -Server $Domaine -Properties * | Sort-Object 
         
        $usersad | ForEach {
 
                  If($_.Surname -eq $null) {$ListViewADuserITEM = New-Object System.Windows.Forms.ListViewItem("")}
                  else {$ListViewADuserITEM = New-Object System.Windows.Forms.ListViewItem($_.Surname)}
                   
                  If($_.GivenName -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.GivenName)}
                  
                  If($_.uidnumber -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.uidnumber)}

                  If($_.EmployeeID -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.EmployeeID)}
                              
                  If($_.ipphone -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.ipphone)}

                  If($_.TelephoneNumber -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.TelephoneNumber)}   
                               
                  If($_.SamAccountName -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.SamAccountName)}               

                  If($_.AccountExpires -eq 0){$ListViewADuserITEM.SubItems.Add("")}             
                  else {$ListViewADuserITEM.SubItems.Add("$(Get-Date([datetime]::FromFileTime($_.AccountExpires)) -Format "d MMM yyyy")")}

                  If($_.Description -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.Description)}                    

                  If($_.mail -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.mail)}  
       
                  $ListViewADuser.Items.Add($ListViewADuserITEM)       

                   }

}

Function LOCKEDADMIN_PROF_PARTEN {         


        $Lockedadmin = Search-ADAccount -AccountExpired -Server $Domaine | Where {($_.ObjectClass -eq 'user')} | Get-ADUser -Properties * | Sort-Object 
     
        $Lockedadmin | ForEach {
 
                  If($_.Surname -eq $null) {$ListViewADuserITEM = New-Object System.Windows.Forms.ListViewItem("")}
                  else {$ListViewADuserITEM = New-Object System.Windows.Forms.ListViewItem($_.Surname)}
                   
                  If($_.GivenName -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.GivenName)}

                  If($_.uidnumber -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.uidnumber)}

                  If($_.EmployeeID -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.EmployeeID)}
           
                  If($_.ipphone -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.ipphone)}

                  If($_.TelephoneNumber -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.TelephoneNumber)}   
                               
                  If($_.SamAccountName -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.SamAccountName)}               

                  If($_.AccountExpires -eq 0){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add("$(Get-Date([datetime]::FromFileTime($_.AccountExpires)) -Format "d MMM yyyy")")}
 
                  If($_.Description -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.Description)}                    

                  If($_.Description -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.mail)}  

                  $ListViewADuser.Items.Add($ListViewADuserITEM)
       
                   }

}


Function Etudiants {
#FENETRE ETUDIANTS
$formEtudiants = New-Object Windows.Forms.Form
$formEtudiants.Add_KeyDown({ if($_.KeyCode -eq "Escape") { $formEtudiants.Close() } }) 
$formEtudiants.KeyPreview = $true
$formEtudiants.Height = 600
$formEtudiants.Width = 800
$formEtudiants.FormBorderStyle = "FixedDialog"
$formEtudiants.MaximizeBox = $false
$formEtudiants.MinimizeBox = $false
$formEtudiants.StartPosition = "CenterScreen"
$formEtudiants.FormBorderStyle = 'Fixed3D' 
$formEtudiants.Text = "personnel"
$formEtudiants.ShowDialog()

}

Function BARREDERECHERCHE {

   If($TextboxADUuser12.Text -eq ""){
    [System.Windows.Forms.MessageBox]::Show("Aucun utilisateur à rechercher.`nVeuillez saisir le 'nom' ou le 'prénom' puis cliquer sur 'Go'","ERREUR RECHERCHE",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Error) 
}  

   
   elseif($TextboxADUuser12.Text -eq $TextboxADUuser12.Text){
       
        $usersad = Get-ADUser -Filter {surname -eq $TextboxADUuser12.Text -or GivenName -eq $TextboxADUuser12.Text} -server $Domaine -Properties *
         
        $usersad | ForEach {
       
                  If($_.Surname -eq $null) {$ListViewADuserITEM = New-Object System.Windows.Forms.ListViewItem("")}
                  else {$ListViewADuserITEM = New-Object System.Windows.Forms.ListViewItem($_.Surname)}
                   
                  If($_.GivenName -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.GivenName)}

                  If($_.uidnumber -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.uidnumber)}

                  If($_.EmployeeID -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.EmployeeID)}
                
                  If($_.ipphone -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.ipphone)}

                  If($_.TelephoneNumber -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.TelephoneNumber)}   
                               
                  If($_.SamAccountName -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.SamAccountName)}               

                  If($_.AccountExpires -eq 0){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add("$(Get-Date([datetime]::FromFileTime($_.AccountExpires)) -Format "dd MMM yyyy")")}

                  If($_.Description -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.Description)}  
                  
                  If($_.Description -eq $null){$ListViewADuserITEM.SubItems.Add("")}
                  else {$ListViewADuserITEM.SubItems.Add($_.mail)}  

                  $ListViewADuser.Items.Add($ListViewADuserITEM)
           
                   }}

}

Function MATRICULE {
  
   If($TextboxModifMatricule1.Text -eq ""){
    [System.Windows.Forms.MessageBox]::Show("Aucun utilisateur séléctionné`nVeuillez relancer votre recherche","ERREUR MATRICULE",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Error) 
} 
  
   elseif($TextboxModifMatricule2.Text -eq ""){           
    [System.Windows.Forms.MessageBox]::Show("Le champ 'numéro de matricule' ne peut pas rester vide.`nVeuillez-vous reporter sur la fiche utilisateur dans Synergie","ERREUR MATRICULE",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Error)

}  

   elseif($TextboxModifMatricule2.Text.Length -lt 5){
    [System.Windows.Forms.MessageBox]::Show("Le matricule doit obigatoirement contenir 5 caractères.`nVeuillez-vous reporter sur la fiche utilisateur dans Synergie","ERREUR MATRICULE",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Error) 
}


   elseif($TextboxModifMatricule2.Text -notmatch "^\d+$"){
    [System.Windows.Forms.MessageBox]::Show("Le matricule doit obigatoirement contenir que des chiffres.`nVeuillez-vous reporter sur la fiche utilisateur dans Synergie","ERREUR MATRICULE",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Error) 
}

#VERIFICATION VALEUR
  elseif((Get-ADUser -filter * -Properties EmployeeID | Select-Object EmployeeID).EmployeeID -eq $($TextboxModifMatricule2.Text)){    
    [System.Windows.Forms.MessageBox]::Show("Ce numéro de matricule est déjà affecté à un autre utilisateur.`nVeuillez-vous reporter sur la fiche utilisateur","ERREUR MATRICULE",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Error)
}
   elseif($TextboxModifMatricule2.Text -match "^\d+$"){
   Set-ADUser -Identity "$($TextboxModifMatricule1.Text)" -EmployeeID "$($TextboxModifMatricule2.Text)" -WhatIf
    [System.Windows.Forms.MessageBox]::Show("Le matricule a été modifié avec succès.`nPour visualiser le changement, relancer la recherche de votre utilisateur.","SUCCES MATRICULE",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information) 

}}

Function MODIFSDA {
  
  If($TextboxModifSDA1.Text -eq ""){
    [System.Windows.Forms.MessageBox]::Show("Aucun utilisateur séléctionné.`nVeuillez relancer votre recherche.","ERREUR SDA",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Error) 

}
   
  elseif($TextboxModifSDA2.Text -eq ""){
    Set-ADUser -Identity "$($TextboxModifSDA1.Text)" -OfficePhone $null   
    [System.Windows.Forms.MessageBox]::Show("Le SDA de l'utilisateur est désormais vide.`nPour visualiser le changement, relancer la recherche de votre utilisateur.","SUCCES SDA",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information)

}

  elseif($TextboxModifSDA2.Text -notmatch "^\d+$"){
    [System.Windows.Forms.MessageBox]::Show("Impossible d'accueillir d'autres caractères à part des chiffres ou vide.`nVeuillez-vous reporter sur le Manager d'Avaya.","ERREUR SDA",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Error)
}

#VERIFICATION VALEUR
  elseif((Get-ADUser -filter * -Properties OfficePhone | Select-Object OfficePhone).OfficePhone -eq $($TextboxModifSDA2.Text)){    
    [System.Windows.Forms.MessageBox]::Show("Il existe déja une SDA affecté à un utilsateur.","ERREUR SDA",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Error)
}
  elseif($TextboxModifSDA2.Text -match "^\d+$"){
    Set-ADUser -Identity "$($TextboxModifSDA1.Text)" -OfficePhone "$($TextboxModifSDA2.Text)"    
    [System.Windows.Forms.MessageBox]::Show("Le numéro de téléphone a été modifié avec succès.`nPour visualiser le changement, relancer la recherche de votre utilisateur.","SUCCES SDA",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information)
}}

Function MODIFEXTENSION {

  If($TextboxModifEXTENSION1.Text -eq ""){
    [System.Windows.Forms.MessageBox]::Show("Aucun utilisateur séléctionné.`nVeuillez relancer votre recherche","ERREUR EXTENSION",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Error) 
}
   
  elseif($TextboxModifEXTENSION2.Text -eq ""){
    Set-ADUser -Identity "$($TextboxModifEXTENSION1.Text)" -Clear ipPhone 
    [System.Windows.Forms.MessageBox]::Show("L'extension de l'utilisateur est désormais vide.`nPour visualiser le changement, relancer la recherche de votre utilisateur.","SUCCES EXTENSION",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information)

}

  elseif($TextboxModifEXTENSION2.Text -notmatch "^\d+$"){  
    [System.Windows.Forms.MessageBox]::Show("Impossible d'accueillir d'autres caractères à part des chiffres ou vide.`nVeuillez-vous reporter sur le Manager d'Avaya.","ERREUR EXTENSION",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Error)

}

#######################################################
#VERIFICATION VALEUR
  elseif((Get-ADUser -filter * -Properties ipphone | Select-Object ipphone).ipphone -eq $($TextboxModifEXTENSION2.Text)){    
    [System.Windows.Forms.MessageBox]::Show("Il existe déja une extension affecté à un utilisateur.","ERREUR EXTENSION",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Error)
}
#######################################################

  elseif($TextboxModifEXTENSION2.Text -match "^\d+$"){  
    Set-ADUser -Identity "$($TextboxModifEXTENSION1.Text)" -Replace @{ipPhone="$($TextboxModifEXTENSION2.Text)"}  
    [System.Windows.Forms.MessageBox]::Show("L'extension a été modifié avec succès.`nPour visualiser le changement, relancer la recherche de votre utilisateur.","SUCCES EXTENSION",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information)

}}

#############################################    

Function COMPTEEXPIRATION {
  If($ModifDateExpiration1.Text -eq ""){
    [System.Windows.Forms.MessageBox]::Show("Aucun utilisateur séléctionné.`nVeuillez relancer votre recherche","Modification date d'expiration",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Error) 
  }
}

Function ACTIVATIONCOMPTE {

  If($ModifDateExpiration1.Text -eq ""){
    [System.Windows.Forms.MessageBox]::Show("Aucun utilisateur séléctionné.`nVeuillez relancer votre recherche"," ERREUR - Activation d'un compte",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Error) 
    }

  Elseif($ModifDateExpiration1.Text -eq $ModifDateExpiration1.Text){ 
    Enable-ADAccount -Identity $ModifDateExpiration1.Text 
    [System.Windows.Forms.MessageBox]::Show("Le compte a été activé avec succès.`nPour visualiser le changement, relancer la recherche de votre utilisateur.","SUCCES - Activation d'un compte",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information) 
    }
 }

Function DESACTIVATIONCOMPTE {

  If($ModifDateExpiration1.Text -eq ""){
    [System.Windows.Forms.MessageBox]::Show("Aucun utilisateur séléctionné.`nVeuillez relancer votre recherche","ERREUR - Désactivation d'un compte",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Error) 
  }

  elseif($ModifDateExpiration1.Text -eq $ModifDateExpiration1.Text){   
    Disable-ADAccount -Identity $ModifDateExpiration1.Text   
    [System.Windows.Forms.MessageBox]::Show("Le compte a été désactivé avec succès.`nPour visualiser le changement, relancer la recherche de votre utilisateur.","SUCCES - Désactivation d'un compte",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information) 
  }
}

Function AJOUTBAL {
  [System.Windows.Forms.MessageBox]::Show("Bientôt disponible...","Ajouter Boite mails",
  [System.Windows.Forms.MessageBoxButtons]::OK,
  [System.Windows.Forms.MessageBoxIcon]::Information) 
}

Function AJOUTVPN {
  [System.Windows.Forms.MessageBox]::Show("Bientôt disponible...","Ajouter VPN",
  [System.Windows.Forms.MessageBoxButtons]::OK,
  [System.Windows.Forms.MessageBoxIcon]::Information) 
}

#endregion Functions
####################################################################################################################################################################### 
                                                              #AFFICHAGE INFORMATIONS USERS TEXTBOX #
####################################################################################################################################################################### 

#region Affichage Informations TEXTBOX

    #NAME
    $ListViewADuser.add_Click({$TextboxADUuser1.text = $ListViewADuser.SelectedItems[0].SubItems[0].Text})
   
    #PRENOM
    $ListViewADuser.add_Click({$TextboxADUuser2.text = $ListViewADuser.SelectedItems[0].SubItems[1].Text})

    #NUMERO CATHO
    $ListViewADuser.add_Click({$TextboxADUuser3.text = $ListViewADuser.SelectedItems[0].SubItems[2].Text})

    #NUMERO MATRICULE
    $ListViewADuser.add_Click({$TextboxADUuser4.text = $ListViewADuser.SelectedItems[0].SubItems[3].Text}) 
                                         
    #IPPHONE (EXTENSION)
    $ListViewADuser.add_Click({$TextboxADUuser5.text = $ListViewADuser.SelectedItems[0].SubItems[4].Text})

    #TELEPHONE (SDA)
    $ListViewADuser.add_Click({$TextboxADUuser6.text = $ListViewADuser.SelectedItems[0].SubItems[5].Text})

    #LOGIN
    $ListViewADuser.add_Click({$TextboxADUuser7.text = $ListViewADuser.SelectedItems[0].SubItems[6].Text})
    $ListViewADuser.add_Click({$TextboxModifMatricule1.text = $ListViewADuser.SelectedItems[0].SubItems[6].Text})
    $ListViewADuser.add_Click({$TextboxModifSDA1.text = $ListViewADuser.SelectedItems[0].SubItems[6].Text})
    $ListViewADuser.add_Click({$TextboxModifEXTENSION1.text = $ListViewADuser.SelectedItems[0].SubItems[6].Text})
    $ListViewADuser.add_Click({$ModifDateExpiration1.text = $ListViewADuser.SelectedItems[0].SubItems[6].Text})

    #DATE EXPIRATION COMPTE
    $ListViewADuser.add_Click({$TextboxADUuser8.text = $ListViewADuser.SelectedItems[0].SubItems[7].Text})
   
    #DESCRIPTION
    $ListViewADuser.add_Click({$TextboxADUuser9.text = $ListViewADuser.SelectedItems[0].SubItems[8].Text})

    #MAIL 10
    $ListViewADuser.add_Click({$TextboxADUuser10.text = $ListViewADuser.SelectedItems[0].SubItems[9].Text})

    #GROUPES TEXTBOX 1
    $ListViewADuser.add_Click({
    $TextboxADUuser11.Text = [string]::join("`r`n",((Get-ADUser -identity ($TextboxADUuser7.Text) -Properties memberof | Select-Object memberof).memberof -replace 'CN=(.+?),(OU|DC)=.+','$1' | sort))   
    }) 
    
    #COMPTE "ETAT COMPTE"
    $ListViewADuser.add_Click({
    $test = (Get-aduser -Identity ($TextboxADUuser7.Text) -Server $Domaine -Properties Enabled | Select-Object Enabled).enabled

    if ($test -match $true){
    $TextboxADUuser13.Text = "Activé"
    }
    elseif ($test -match $False){
    $TextboxADUuser13.Text = "Désactivé"
    }})


####################################################################################################################################################################### 
    #AFFICHE CONTENU TEXTBOX
    $formADuser.controls.Add($ListViewADuser)

#endregion Affichage Informations TEXTBOX

####################################################################################################################################################################### 
                                                                     #AFFICHAGE CONTENU GENERAL INFORMATIONS USERS
####################################################################################################################################################################### 
    $formADuser.ShowDialog()
####################################################################################################################################################################### 
}
#endregion
#######################################################################################################################################################################
                                                                          # FENETRES TRANSFERT DE FICHIER #
#######################################################################################################################################################################
#region VERIFICATION NUMERO

function TEL {

#FENETRE TEL
$FORM_TEL = New-Object Windows.Forms.Form
$FORM_TEL.Add_KeyDown({if($_.KeyCode -eq "Escape"){$FORM_TEL.Close()}}) 
$FORM_TEL.KeyPreview = $true
$FORM_TEL.Height = 250
$FORM_TEL.Width = 500
$FORM_TEL.FormBorderStyle = "FixedDialog"
$FORM_TEL.MaximizeBox = $false
$FORM_TEL.MinimizeBox = $false
$FORM_TEL.StartPosition = "CenterScreen"
$FORM_TEL.FormBorderStyle = 'Fixed3D'
$FORM_TEL.Text = "VERIFICATION NUMERO"

#Label1 TEL
Add-Type -AssemblyName System.Windows.Forms
$TEL_LABEL1 = New-Object System.Windows.Forms.Label
$TEL_LABEL1.Location = New-Object System.Drawing.Size(150,20)
$TEL_LABEL1.Size = New-Object System.Drawing.Size(275,15)
$TEL_LABEL1.Text = "VERIFICATION NUMERO"
$TEL_LABEL1.Font = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Bold)
$TEL_LABEL1.AutoSize = $false
$TEL_LABEL1.Add_KeyDown({if($_.KeyCode -eq "Escape"){$FORM_TEL.Close()}}) 
$FORM_TEL.Controls.Add($TEL_LABEL1)

#Label2 TEL
Add-Type -AssemblyName System.Windows.Forms
$TEL_LABEL2 = New-Object System.Windows.Forms.Label
$TEL_LABEL2.Location = New-Object System.Drawing.Size(25,65)
$TEL_LABEL2.Size = New-Object System.Drawing.Size(298,15)
$TEL_LABEL2.Text = "Veuillez saisir l'extension ou le numéro de téléphone :"
$TEL_LABEL2.Font = New-Object System.Drawing.Font("Arial",9,[System.Drawing.FontStyle]::Regular)
$TEL_LABEL2.AutoSize = $false
$TEL_LABEL2.Add_KeyDown({ if($_.KeyCode -eq "Escape"){$FORM_TEL.Close()}}) 
$FORM_TEL.Controls.Add($TEL_LABEL2)

#TEXTBOX SAISI
Add-Type -AssemblyName System.Windows.Forms
$TEL_TEXTBOX1 = New-Object System.Windows.Forms.TextBox 
$TEL_TEXTBOX1.Location = New-Object System.Drawing.Size(30,90)
$TEL_TEXTBOX1.Size = New-Object System.Drawing.Size(190,20) 
$TEL_TEXTBOX1.BackColor = "white"
$TEL_TEXTBOX1.MaxLength = 10
$TEL_TEXTBOX1.Add_KeyDown({if($_.KeyCode -eq "Escape"){$FORM_TEL.Close()}}) 
$FORM_TEL.Controls.Add($TEL_TEXTBOX1)

#BUTTON REQUETE
Add-Type -AssemblyName System.Windows.Forms
$TEL_BUTTON1= New-Object System.Windows.Forms.Button
$TEL_BUTTON1.Text = "GO"
$TEL_BUTTON1.Font = New-Object System.Drawing.Font("Arial",9,[System.Drawing.FontStyle]::BOLD)
$TEL_BUTTON1.Width = 90
$TEL_BUTTON1.Height = 30
$TEL_BUTTON1.Location = New-Object System.Drawing.Size(360,85)
$TEL_BUTTON1.Add_KeyDown({if($_.KeyCode -eq "Escape"){$FORM_TEL.Close()}}) 
$TEL_BUTTON1.add_Click({$TEL_TEXTBOX2.Clear()})
$TEL_BUTTON1.add_Click({TELEPHONE})
$FORM_TEL.Controls.Add($TEL_BUTTON1)

#BUTTON FERMER
Add-Type -AssemblyName System.Windows.Forms
$TEL_BUTTON2= New-Object System.Windows.Forms.Button
$TEL_BUTTON2.Text = "FERMER"
$TEL_BUTTON2.Font = New-Object System.Drawing.Font("Arial",9,[System.Drawing.FontStyle]::BOLD)
$TEL_BUTTON2.Width = 90
$TEL_BUTTON2.Height = 30
$TEL_BUTTON2.Location = New-Object System.Drawing.Size(360,125)
$TEL_BUTTON2.Add_KeyDown({if($_.KeyCode -eq "Escape"){$FORM_TEL.Close()}}) 
$TEL_BUTTON2.add_Click({$FORM_TEL.Close()})
$FORM_TEL.Controls.Add($TEL_BUTTON2)

#BUTTON FERMER PROGRAMME
Add-Type -AssemblyName System.Windows.Forms
$TEL_BUTTON2= New-Object System.Windows.Forms.Button
$TEL_BUTTON2.Text = "FERMER PROGRAMME"
$TEL_BUTTON2.Font = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Bold)
$TEL_BUTTON2.Width = 90
$TEL_BUTTON2.Height = 30
$TEL_BUTTON2.Location = New-Object System.Drawing.Size(360,165)
$TEL_BUTTON2.Add_KeyDown({if($_.KeyCode -eq "Escape"){$FORM_TEL.Close()}}) 
$TEL_BUTTON2.add_Click({$FORM_TEL.Close()})
$TEL_BUTTON2.add_Click({$form.Close()})
$FORM_TEL.Controls.Add($TEL_BUTTON2)

#Label3
Add-Type -AssemblyName System.Windows.Forms
$TEL_LABEL3 = New-Object System.Windows.Forms.Label
$TEL_LABEL3.Location = New-Object System.Drawing.Size(27,130)
$TEL_LABEL3.Size = New-Object System.Drawing.Size(275,15)
$TEL_LABEL3.Text = "Resultat :"
$TEL_LABEL3.AutoSize = $false
$TEL_LABEL3.Add_KeyDown({if($_.KeyCode -eq "Escape"){$FORM_TEL.Close()}}) 
$TEL_LABEL3.Font = New-Object System.Drawing.Font("Arial",9,[System.Drawing.FontStyle]::Regular)
$FORM_TEL.Controls.Add($TEL_LABEL3)

#TEXTBOX REPONSE REQUETE
Add-Type -AssemblyName System.Windows.Forms
$TEL_TEXTBOX2 = New-Object System.Windows.Forms.TextBox 
$TEL_TEXTBOX2.Location = New-Object System.Drawing.Size(30,155)
$TEL_TEXTBOX2.Size = New-Object System.Drawing.Size(300,200) 
$TEL_TEXTBOX2.ReadOnly = $true
$TEL_TEXTBOX2.Add_KeyDown({if($_.KeyCode -eq "Escape"){$FORM_TEL.Close()}}) 
$FORM_TEL.Controls.Add($TEL_TEXTBOX2)
#
######################################################################################################################################################################
                                                                                   # FUNCTIONS #
######################################################################################################################################################################

Function TELEPHONE {

  If($TEL_TEXTBOX1.Text -eq ""){
    [System.Windows.Forms.MessageBox]::Show("Aucun numéro séléctionné.`nVeuillez saisir une extension ou une S.D.A","ERREUR",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Error) 
} 

   elseif($TEL_TEXTBOX1.Text -notmatch "^\d+$"){
    [System.Windows.Forms.MessageBox]::Show("Un numéro de téléphone ne peut pas accepter d'autres caractères à part des chiffres.","ERREUR",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Error) 
}

   elseif($TEL_TEXTBOX1.Text.Length -lt 4){
    [System.Windows.Forms.MessageBox]::Show("Un numéro de téléphone peut contenir entre 4 et 10 chiffres.","ERREUR",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Error) 
}

#VERIFICATION + REQUETE (EXTENSION) 
  elseif((Get-ADUser -filter * -Properties ipphone | Select-Object ipphone).ipphone -eq $($TEL_TEXTBOX1.Text)){    
  $requete1 = Get-ADUser -LDAPFilter "(Ipphone=$($TEL_TEXTBOX1.Text))"
  $TEL_TEXTBOX2.Text = "Ce numéro appartient à $($requete1.Name)"
}

#VERIFICATION + REQUETE (SDA) 
  elseif((Get-ADUser -Filter { OfficePhone -eq $TEL_TEXTBOX1.Text})){    
  $requete2 = (Get-ADUser -Filter { OfficePhone -eq $TEL_TEXTBOX1.Text}) 
  $TEL_TEXTBOX2.Text = "Ce numéro appartient à $($requete2.Name)"
}
#VERIFICATION (EXTENSION)
  elseif((Get-ADUser -filter * -Properties ipphone | Select-Object ipphone).ipphone -cnotmatch $($TEL_TEXTBOX1.Text)){  
    [System.Windows.Forms.MessageBox]::Show("Aucun utilisateur à ce numéro.","ERREUR",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Error)   
}}

######################################################################################################################################################################
                                                                             # AFFICHAGE DE LA FENETRE #
######################################################################################################################################################################
$FORM_TEL.ShowDialog() 
######################################################################################################################################################################
}
#endregion
#######################################################################################################################################################################
                                                                          # FENETRES INSTALLATION LOGICIEL#
#######################################################################################################################################################################
#region INSTALLATION LOGICIEL

function INSTALLSOFTWARE {
  [System.Windows.Forms.MessageBox]::Show("En cours de construction...","Installation d'un logiciel",
  [System.Windows.Forms.MessageBoxButtons]::OK,
  [System.Windows.Forms.MessageBoxIcon]::Information)
}
#endregion
#######################################################################################################################################################################
                                                                          # FENETRES USB BOOTABLE#
#######################################################################################################################################################################
#region USbBOOTABLE

function USbBOOTABLE {
  [System.Windows.Forms.MessageBox]::Show("En cours de construction...","Création d'une clé USB BOOTABLE",
  [System.Windows.Forms.MessageBoxButtons]::OK,
  [System.Windows.Forms.MessageBoxIcon]::Information)
}
#endregion
#######################################################################################################################################################################
                                                                          # FENETRES DEPLACEMENT ORDINATEUR#
#######################################################################################################################################################################
#region DEPLACEMENT ORDINATEUR 

function DeplacementpcDOMAIN {
  [System.Windows.Forms.MessageBox]::Show("En cours de construction...","Déplacer un ordinateur (Active Directory)",
  [System.Windows.Forms.MessageBoxButtons]::OK,
  [System.Windows.Forms.MessageBoxIcon]::Information)
}
#endregion 
######################################################################################################################################################################
                                                                             # BUTTONS FENETRE PRINCIPALE (FORM) #
######################################################################################################################################################################

#region BUTTONS MENU
#BOUTON VALIDER (MENU)
$button_valider = New-Object System.Windows.Forms.Button
$button_valider.Text = "Valider"
$button_valider.Size = New-Object System.Drawing.Size(90,35)
$button_valider.Location = New-Object System.Drawing.Size(259,540)
$button_valider.Font = ‘Arial,13’
$button_valider.BackColor = "Lavender"
$form.Controls.Add($button_valider)

#BOUTON QUITTER (MENU)
$button_quitter = New-Object System.Windows.Forms.Button
$button_quitter.Text = "Quitter"
$button_quitter.Size = New-Object System.Drawing.Size(90,35)
$button_quitter.Location = New-Object System.Drawing.Size(47,540)
$button_quitter.Font = ‘Arial,13’
$button_quitter.BackColor = "Lavender"
$button_quitter.Add_Click({$form.close()})
$form.Controls.Add($button_quitter)

#BOUTON SELECTION (MENU)
$button_valider.add_Click({

    if ($ListBox_1.SelectedIndex -eq 0) {INFORMATIONsUSERS} 
    elseif ($ListBox_1.SelectedIndex -eq 1) {TEL}   
    elseif ($ListBox_1.SelectedIndex -eq 2) {INSTALLSOFTWARE}
    elseif ($ListBox_1.SelectedIndex -eq 3) {USbBOOTABLE}
    elseif ($ListBox_1.SelectedIndex -eq 4) {DeplacementpcDOMAIN} 

})
#endregion
######################################################################################################################################################################
                                                                             # AFFICHAGE DE LA FENETRE #
######################################################################################################################################################################
$form.ShowDialog() | out-null
######################################################################################################################################################################
                                                                                # END OF PROGRAMM #
######################################################################################################################################################################
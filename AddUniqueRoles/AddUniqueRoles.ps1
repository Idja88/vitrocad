#Переменные
$SiteURL = ""
$ListName="Проекты"
$date = Get-Date

#Подключение к вебсайту под sp_setup
Connect-PnPOnline -Url $SiteURL -CurrentCredentials

#Задача запускается раз в 15 минут, поэтому находим все папки созданные за 20 минут до запуска
$query = "<View Scope='RecursiveAll'>
            <Query>
                <Where>
                    <And>
                        <Eq>
                            <FieldRef Name='FSObjType' />
                            <Value Type='Integer'>1</Value>
                        </Eq>
                        <And>
                            <Geq>
                                <FieldRef Name='Created' />
                                <Value IncludeTimeValue='TRUE' Type='DateTime'>" + ($date.AddMinutes(-20)).ToString("s") + "Z</Value>
                            </Geq>
                            <Leq>
                                <FieldRef Name='Created' />
                                <Value IncludeTimeValue='TRUE' Type='DateTime'>" + $date.ToString("s") + "Z</Value>
                            </Leq>
                        </And>
                    </And>
                </Where>
            </Query>
        </View>"

$ListItems = Get-PnPListItem -List $ListName -Query $query

$Folders = $ListItems | Where-Object {$_.FieldValues.VitroProjectViewerRole -ne $null}

    #Pipeline
    $Folders | ForEach-Object {

    Try {

        $Folder = $_.Folder

        #Получаем доп свойства папки
        $FolderItem = Get-PnPProperty -ClientObject $Folder -Property ListItemAllFields
        $HasUniquePerm = Get-PnPProperty -ClientObject $FolderItem -Property HasUniqueRoleAssignments

        #Обрываем наследование прав
        If(!$HasUniquePerm){

            $FolderItem.BreakRoleInheritance($False, $False)
            Write-Host "`tFolder's Permission Inheritance Broken!"

            #Получаем пользователей которые имеют права после обрыва
            $RoleAssignments = Get-PnPProperty -ClientObject $FolderItem -Property RoleAssignments

            #Убираем всех пользователей/группы пользователей что имеют права после обрыва
            ForEach($RoleAssignment in $RoleAssignments)
            {
                $Member =  Get-PnPProperty -ClientObject $RoleAssignment -Property Member

                $FolderItem.RoleAssignments.GetByPrincipal($Member).DeleteObject()
                Invoke-PnPQuery
                Write-Host "`tRemoved $($Member.Title) from Folder Permission!"
            }

            #Объект всех пользователей/групп пользователей в поле VitroProjectViewerRole
            $FolderItemVitroGroups = $Folder.ListItemAllFields.FieldValues.VitroProjectViewerRole

            #Проставляем уникальные разрешения для папок
            ForEach($FolderItemVitroGroup in $FolderItemVitroGroups){
                $group = Get-PnPGroup -Identity $FolderItemVitroGroup.LookupId -ErrorAction SilentlyContinue

                if ($group) {
                    Set-PnPListItemPermission -List $ListName -Identity $Folder.ListItemAllFields -Group $group.LoginName -AddRole 'Ответственный'
                    Write-Host "`tAdded $($group.Title) to Folder Permission!"
                }
                else {
                    Set-PnPListItemPermission -List $ListName -Identity $Folder.ListItemAllFields -User $FolderItemVitroGroup.LookupValue -AddRole 'Ответственный'
                    Write-Host "`tAdded $($FolderItemVitroGroup.LookupValue) to Folder Permission!"
                }
            }
        }

    }
    Catch {
        write-host -f Red "Error Removing all users or groups from Folder:" $_.Exception.Message
    }
}
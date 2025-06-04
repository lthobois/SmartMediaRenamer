$AppVersion = "SmartMediaRenamer v3.0 (2025-05-01)"
$ApiUrl = "https://api.themoviedb.org/3"
$ApiKey = "91daaaa00067aa11822875c46ddb9204"
$ApiLanguage = "fr-FR"
$MediaScanPath = "C:\Users\lthob\Downloads\TestMedia"
$MediaTypeFilter = "*.mp4", "*.mkv", "*.avi"
$MoviePatternOutput = "X:\{CategoryName}\{MovieName} ({MovieYear}).{FileExtension}"
$TvShowPatternOutput = "Y:\{TvShowName}\Saison {SeasonNumber}\{TvShowName} {SeasonNumber}X{EpisodeNumber} - {EpisodeTitle}.{FileExtension}"
$TvShowCacheFilePath = "$env:LOCALAPPDATA\SmartMediaRenamer\TvShowCache.json"

Clear-Host

#region File management

function Format-FileName {
    param (
        [string]$fileName
    )

    $wordsToRemove = @("(", ")", "FRENCH", "1080p", "H264", "FiNAL", "DSNP", "WEB-DL", "DDP5.1", "Wawacity", "moe", "MULTi", "WEBRip", "x265", "-KFL", "VFI", "10bit", "HDR", "DDP", "5.1", "-ASKO", "SOCIAL", "BIKE", "Zone", "Telechargement", "[OFFICIAL]", "MUTE", "HDLight", "AC3", "LiHDL", "True", "com", "mp4", "avi", "mkv" )
    $cleanFileName = $fileName

    foreach ($word in $wordsToRemove) {
        $cleanFileName = $cleanFileName -replace [regex]::Escape($word), ""
    }

    # Supprime espaces multiple
    $cleanFileName = $cleanFileName -replace '\s+', ' '
    # Supprime les années
    $cleanFileName = $cleanFileName -replace '\b\d{4}\b', ''
    # Supprime les points et tirets
    $cleanFileName = $cleanFileName -replace '\.|-', ' '

    return $cleanFileName.Trim()
}

function Get-MediaFiles {
    if (-not (Test-Path $global:currentFolder)) { return }
    $global:fileList = Get-ChildItem -Path $global:currentFolder -Recurse -File -Include $MediaTypeFilter
    $global:fileData = $global:fileList | ForEach-Object {
        [PSCustomObject]@{
            IsSelected = $true
            Type = Get-MediaType -fileName $_.Name
            OldName = $_.FullName
            NewName = ""
        }
    }
    Set-Filter
}

function Get-MediaType {
    param ( [string]$fileName )

    # Vérifier si le nom de fichier contient des informations de saison ou d'épisode
    if (($fileName -match "S\d{2}E\d{2}" -or $fileName -match "S\d{2}") -and $fileName -match "\.mp4$|\.mkv$|\.avi$") {
        return "TvShow"
    } elseif ($fileName -match "\.mp4$|\.mkv$|\.avi$") {
        return "Movie"
    } else {
        return ""
    }
}

function Set-Filter {
    $filterText = $txtFilter.Text.ToLower()
    if ([string]::IsNullOrWhiteSpace($filterText)) {
        $global:filteredData = $global:fileData
    } else {
        $global:filteredData = $global:fileData | Where-Object { $_.OldName.ToLower().Contains($filterText) }
    }
    $fileGrid.ItemsSource = $null
    $fileGrid.ItemsSource = $global:filteredData
}

#endregion

#region Movie

function Get-MovieItemFromMovieDB {
    param (
        [string]$fileName
    )

    $cleanFileName = Format-FileName -fileName $fileName
    Write-Host "Get-MovieItemFromMovieDB : $cleanFileName" -ForegroundColor Green

    $encodedQuery = [uri]::EscapeDataString($cleanFileName)
    $searchUrl = "$ApiUrl/search/movie?api_key=$ApiKey&query=$encodedQuery&language=$ApiLanguage"

    try {
        $movieResponse = Invoke-RestMethod -Uri $searchUrl -Method Get
    } catch {
        Write-Host "Erreur lors de l'appel à l'API TMDb : $_" -ForegroundColor Red
        return $null
    }

    if (-not $movieResponse.results -or $movieResponse.total_results -eq 0) {
        Write-Host "Aucun résultat trouvé pour : $cleanFileName"
        return $null
    }

    $movieList = $movieResponse.results

    if ($movieList.Count -eq 1) {
        $selection = $movieList[0]
    } else {
        Add-Type -AssemblyName PresentationFramework

        # Créer la fenêtre de sélection
        $selectionWindow = New-Object System.Windows.Window
        $selectionWindow.Title = "$($fileName): Sélectionnez le film"
        $selectionWindow.WindowStyle = [System.Windows.WindowStyle]::SingleBorderWindow
        $selectionWindow.WindowState = [System.Windows.WindowState]::Maximized
        $selectionWindow.WindowStartupLocation = "CenterScreen"

        # Créer un Grid pour organiser les éléments
        $grid = New-Object System.Windows.Controls.Grid
        $grid.Margin = "10"

        # Définir les colonnes du Grid
        $col1 = New-Object System.Windows.Controls.ColumnDefinition
        $col2 = New-Object System.Windows.Controls.ColumnDefinition
        $col1.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
        $col2.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
        $grid.ColumnDefinitions.Add($col1)
        $grid.ColumnDefinitions.Add($col2)

        # Créer une ListBox pour afficher les films
        $listBox = New-Object System.Windows.Controls.ListBox
        $listBox.Margin = "0,0,10,0"
        $listBox.HorizontalAlignment = "Stretch"
        $listBox.VerticalAlignment = "Stretch"
        [System.Windows.Controls.Grid]::SetColumn($listBox, 0)
        $grid.Children.Add($listBox)

        foreach ($movie in $movieList) {
            $title = if ($movie.title) { $movie.title } else { "[Titre inconnu]" }
            $date  = if ($movie.release_date) { $movie.release_date } else { "???" }
            $listBox.Items.Add("$title ($date)")
        }

        # Créer un contrôle Image pour afficher la jaquette
        $image = New-Object System.Windows.Controls.Image
        $image.Stretch = [System.Windows.Media.Stretch]::Uniform
        $image.HorizontalAlignment = "Stretch"
        $image.VerticalAlignment = "Stretch"
        [System.Windows.Controls.Grid]::SetColumn($image, 1)
        $grid.Children.Add($image)

        # Ajouter le Grid à la fenêtre
        $selectionWindow.Content = $grid

        # Variable globale pour stocker l'index du film sélectionné
        $global:selectedMovieIndex = -1

        # Fonction pour mettre à jour l'image
        function Update-Image {
            param ($movie)

            $posterUrl = "https://image.tmdb.org/t/p/original$($movie.poster_path)"

            # Créer une source d'image à partir de l'URL
            $bitmapImage = New-Object System.Windows.Media.Imaging.BitmapImage
            $bitmapImage.BeginInit()
            $bitmapImage.UriSource = $posterUrl
            $bitmapImage.EndInit()

            # Définir la source de l'image
            $image.Source = $bitmapImage
        }

        # Sélectionner le premier film par défaut et afficher son image
        if ($listBox.Items.Count -gt 0) {
            $listBox.SelectedIndex = 0
            Update-Image -movie $movieList[0]
        }

        # Ajouter un gestionnaire d'événements pour la sélection dans la ListBox
        $listBox.Add_SelectionChanged({
            if ($listBox.SelectedIndex -ge 0) {
                Update-Image -movie $movieList[$listBox.SelectedIndex]
            }
        })

        # Ajouter un gestionnaire d'événements pour le double-clic sur un élément de la ListBox
        $listBox.Add_MouseDoubleClick({
            if ($listBox.SelectedIndex -ge 0) {
                $global:selectedMovieIndex = $listBox.SelectedIndex
                $selectionWindow.Close()
            }
        })

        # Ajouter un gestionnaire d'événements pour la touche Entrée
        $selectionWindow.Add_KeyDown({
            if ($_.Key -eq [System.Windows.Input.Key]::Enter) {
                if ($listBox.SelectedIndex -ge 0) {
                    $global:selectedMovieIndex = $listBox.SelectedIndex
                    $selectionWindow.Close()
                }
            }
        })

        # Ajouter un gestionnaire d'événements pour donner le focus à la ListBox lorsque la fenêtre est chargée
        $selectionWindow.Add_Loaded({
            $listBox.Focus()
        })

        # Afficher la fenêtre
        $selectionWindow.ShowDialog() | Out-Null

        if ($global:selectedMovieIndex -ge 0) {
            $selection = $movieList[$global:selectedMovieIndex]
        } else {
            Write-Host "Aucune sélection effectuée."
            return $null
        }
    }

    $values = @{
        FileName = $fileName
        FileExtension = [System.IO.Path]::GetExtension($fileName).TrimStart('.')
        MovieId = $selection.id
        MovieName = $selection.title
        MovieYear = $selection.release_date.Substring(0,4)
        MoviePoster = "https://image.tmdb.org/t/p/original$($selection.poster_path)"
        CategoryName = Get-MovieCategoryFromMovieDB -Id $selection.genre_ids[0]
        GenreIds = $selection.genre_ids
    }

    $formattedPath = $MoviePatternOutput
    $values.GetEnumerator() | ForEach-Object {
        $formattedPath = $formattedPath -replace "{$($_.Name)}", $_.Value
    }

    $formattedPath = "$($formattedPath[0..1])$($formattedPath.Substring(2) -replace ":", " -")".Replace(" :\",":\")
    $formattedPath = $formattedPath -replace "[$([regex]::Escape([IO.Path]::GetInvalidPathChars() -join ''))]", ""
    return $formattedPath.Replace('?', '')
}

function Get-MovieCategoryFromMovieDB {
    param (
        [int]$Id
    )

    # Créer un tableau des valeurs id / genre
    $genreTable = @(
        [PSCustomObject]@{ ID = 28; Genre = "Action" },
        [PSCustomObject]@{ ID = 12; Genre = "Aventure" },
        [PSCustomObject]@{ ID = 16; Genre = "Animation" },
        [PSCustomObject]@{ ID = 35; Genre = "Comédie" },
        [PSCustomObject]@{ ID = 80; Genre = "Crime" },
        [PSCustomObject]@{ ID = 99; Genre = "Documentaire" },
        [PSCustomObject]@{ ID = 18; Genre = "Drame" },
        [PSCustomObject]@{ ID = 10751; Genre = "Familial" },
        [PSCustomObject]@{ ID = 14; Genre = "Fantastique" },
        [PSCustomObject]@{ ID = 36; Genre = "Histoire" },
        [PSCustomObject]@{ ID = 27; Genre = "Horreur" },
        [PSCustomObject]@{ ID = 10402; Genre = "Musique" },
        [PSCustomObject]@{ ID = 9648; Genre = "Mystère" },
        [PSCustomObject]@{ ID = 10749; Genre = "Romance" },
        [PSCustomObject]@{ ID = 878; Genre = "Science-Fiction" },
        [PSCustomObject]@{ ID = 10770; Genre = "Téléfilm" },
        [PSCustomObject]@{ ID = 53; Genre = "Thriller" },
        [PSCustomObject]@{ ID = 10752; Genre = "Guerre" },
        [PSCustomObject]@{ ID = 37; Genre = "Western" }
    )

    # Rechercher le genre correspondant à l'ID
    $genre = $genreTable | Where-Object { $_.ID -eq $Id }

    if ($genre) {
        return $genre.Genre
    } else {
        Write-Host "Catégorie non trouvé."
        return "Divers"
    }
}

#endregion

#region TvShow

# Charge ou crée un Hashtable pour le cache
if (Test-Path $TvShowCacheFilePath) {
    $json = Get-Content -Raw -Encoding UTF8 -Path $TvShowCacheFilePath
    $cacheObj = $json | ConvertFrom-Json
    $TvShowCache = @{ }
    # Conversion PSCustomObject -> Hashtable
    $cacheObj.PSObject.Properties | ForEach-Object {
        $TvShowCache[$_.Name] = $_.Value
    }
} else {
    $TvShowCache = @{ }
}

function Get-TvShowItemFromMovieDB {
    param (
        [string]$TvShowName
    )

    Write-Host "Get-TvShowItemFromMovieDB : $TvShowName" -ForegroundColor Green

    # Nettoyer la chaîne pour la clé du cache
    $key = $TvShowName.Trim().ToLower()

    # Si présent dans le cache, on retourne directement
    if ($TvShowCache.ContainsKey($key)) {
        $cached = $TvShowCache[$key]
        Write-Host "Série trouvée en cache : $($cached.name) (ID $($cached.id))"
        return $cached
    }

    try {
        $searchUrl = "$ApiUrl/search/tv?api_key=$ApiKey&query=$TvShowName&language=$ApiLanguage".Replace(" ", "%20")
        $seriesResponse = Invoke-RestMethod -Uri $searchUrl
    } catch {
        Write-Host "Erreur lors de l'appel à l'API TMDb : $_" -ForegroundColor Red
        return $null
    }

    $seriesList = $seriesResponse.results

    if ($seriesList.Count -eq 0) {
        return $null
    } elseif ($seriesList.Count -eq 1) {
        $selection = $seriesList[0]
    } else {
        Add-Type -AssemblyName PresentationFramework

        # Créer la fenêtre de sélection
        $selectionWindow = New-Object System.Windows.Window
        $selectionWindow.Title = "$($TvShowName): Sélectionnez la série"
        $selectionWindow.WindowStyle = [System.Windows.WindowStyle]::SingleBorderWindow
        $selectionWindow.WindowState = [System.Windows.WindowState]::Maximized
        $selectionWindow.WindowStartupLocation = "CenterScreen"

        # Créer un Grid pour organiser les éléments
        $grid = New-Object System.Windows.Controls.Grid
        $grid.Margin = "10"

        # Définir les colonnes du Grid
        $col1 = New-Object System.Windows.Controls.ColumnDefinition
        $col2 = New-Object System.Windows.Controls.ColumnDefinition
        $col1.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
        $col2.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
        $grid.ColumnDefinitions.Add($col1)
        $grid.ColumnDefinitions.Add($col2)

        # Créer une ListBox pour afficher les séries
        $listBox = New-Object System.Windows.Controls.ListBox
        $listBox.Margin = "0,0,10,0"
        $listBox.HorizontalAlignment = "Stretch"
        $listBox.VerticalAlignment = "Stretch"
        [System.Windows.Controls.Grid]::SetColumn($listBox, 0)
        $grid.Children.Add($listBox)

        foreach ($series in $seriesList) {
            $name = if ($series.name) { $series.name } else { "[Nom inconnu]" }
            $date  = if ($series.first_air_date) { $series.first_air_date } else { "???" }
            $listBox.Items.Add("$name ($date)")
        }

        # Créer un contrôle Image pour afficher la jaquette
        $image = New-Object System.Windows.Controls.Image
        $image.Stretch = [System.Windows.Media.Stretch]::Uniform
        $image.HorizontalAlignment = "Stretch"
        $image.VerticalAlignment = "Stretch"
        [System.Windows.Controls.Grid]::SetColumn($image, 1)
        $grid.Children.Add($image)

        # Ajouter le Grid à la fenêtre
        $selectionWindow.Content = $grid

        # Variable globale pour stocker l'index de la série sélectionnée
        $global:selectedSeriesIndex = -1

        # Fonction pour mettre à jour l'image
        function Update-Image {
            param ($series)

            $posterUrl = "https://image.tmdb.org/t/p/original$($series.poster_path)"

            # Créer une source d'image à partir de l'URL
            $bitmapImage = New-Object System.Windows.Media.Imaging.BitmapImage
            $bitmapImage.BeginInit()
            $bitmapImage.UriSource = $posterUrl
            $bitmapImage.EndInit()

            # Définir la source de l'image
            $image.Source = $bitmapImage
        }

        # Sélectionner la première série par défaut et afficher son image
        if ($listBox.Items.Count -gt 0) {
            $listBox.SelectedIndex = 0
            Update-Image -series $seriesList[0]
        }

        # Ajouter un gestionnaire d'événements pour la sélection dans la ListBox
        $listBox.Add_SelectionChanged({
            if ($listBox.SelectedIndex -ge 0) {
                Update-Image -series $seriesList[$listBox.SelectedIndex]
            }
        })

        # Ajouter un gestionnaire d'événements pour le double-clic sur un élément de la ListBox
        $listBox.Add_MouseDoubleClick({
            if ($listBox.SelectedIndex -ge 0) {
                $global:selectedSeriesIndex = $listBox.SelectedIndex
                $selectionWindow.Close()
            }
        })

        # Ajouter un gestionnaire d'événements pour la touche Entrée
        $selectionWindow.Add_KeyDown({
            if ($_.Key -eq [System.Windows.Input.Key]::Enter) {
                if ($listBox.SelectedIndex -ge 0) {
                    $global:selectedSeriesIndex = $listBox.SelectedIndex
                    $selectionWindow.Close()
                }
            }
        })

        # Ajouter un gestionnaire d'événements pour donner le focus à la ListBox lorsque la fenêtre est chargée
        $selectionWindow.Add_Loaded({
            $listBox.Focus()
        })

        # Afficher la fenêtre
        $selectionWindow.ShowDialog() | Out-Null

        if ($global:selectedSeriesIndex -ge 0) {
            $selection = $seriesList[$global:selectedSeriesIndex]
        } else {
            Write-Host "Aucune sélection effectuée."
            return $null
        }
    }

    # Sauvegarde dans le cache
    $TvShowCache[$key] = $selection
    Save-TvShowCache -cache $TvShowCache -path $TvShowCacheFilePath

    return $selection
}

function Get-TvShowEpisodeTitleFromMovieDB {
    param (
        [string]$fileName
    )

    $cleanFileName = Format-FileName -fileName $fileName
    Write-Host "Get-TvShowEpisodeTitleFromMovieDB : $fileName" -ForegroundColor Green

    if ($cleanFileName -match "(.*)S(\d{2})E(\d{2})(.*)") {
        $TvShowName = $matches[1].Trim()
        $seasonNumber = [int]$matches[2]
        $episodeNumber = [int]$matches[3]
        $fileExtension = [System.IO.Path]::GetExtension($fileName).TrimStart('.')

        $seriesItem = Get-TvShowItemFromMovieDB -TvShowName $TvShowName

        if ($seriesItem -and $seriesItem.id) {
            try {
                $episodeUrl = "$ApiUrl/tv/$($seriesItem.id)/season/$seasonNumber/episode/$($episodeNumber)?api_key=$ApiKey&language=$ApiLanguage"
                $episodeResponse = Invoke-RestMethod -Uri $episodeUrl
            } catch {
                Write-Host "Erreur lors de l'appel à l'API TMDb : $_" -ForegroundColor Red
                return $null
            }

            $episodeTitle = $episodeResponse.name

            $values = @{
                FileName = $fileName
                FileExtension = $fileExtension
                TvShowId = $seriesItem.id
                TvShowName = $seriesItem.name  # Utilisation du nom en français
                SeasonNumber = $seasonNumber.ToString("D2")
                EpisodeNumber = $episodeNumber.ToString("D2")
                EpisodeTitle = $episodeTitle
            }

            $formattedPath = $TvShowPatternOutput
            $values.GetEnumerator() | ForEach-Object {
                $formattedPath = $formattedPath -replace "{$($_.Name)}", $_.Value
            }

            $formattedPath = "$($formattedPath[0..1])$($formattedPath.Substring(2) -replace ":", " -")".Replace(" :\",":\")
            $formattedPath = $formattedPath -replace "[$([regex]::Escape([IO.Path]::GetInvalidPathChars() -join ''))]", ""

            return $formattedPath.Replace('?','')
        }
    } else {
        return "Format de nom de fichier incorrect."
    }
}

function Save-TvShowCache {
    param($cache, $path)
    # Crée le dossier si besoin
    $dir = Split-Path $path -Parent
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $cache | ConvertTo-Json -Depth 4 | Set-Content -Path $path -Encoding UTF8
}

#endregion

#region Windows

Add-Type -AssemblyName PresentationFramework

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="$AppVersion" Height="600" Width="800" WindowState="Maximized">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
            <Button Name="btnSelectFolder" Content="Choisir Répertoire" Width="150" Margin="0,0,10,0"/>
            <Button Name="btnRefresh" Content="Rafraîchir" Width="100" Margin="0,0,10,0"/>
            <Button Name="btnGenerateNames" Content="Générer les noms" Width="150" Margin="0,0,10,0"/>
            <Button Name="btnRename" Content="Renommer" Width="100" Margin="0,0,10,0"/>
            <CheckBox Name="chkSelectAll" Content="Tout sélectionner / désélectionner" Margin="10,0,0,0"/>
        </StackPanel>
        <StackPanel Orientation="Horizontal" Grid.Row="1" Margin="0,0,0,10">
            <TextBlock Text="Filtrer : " VerticalAlignment="Center" Margin="0,0,5,0"/>
            <TextBox Name="txtFilter" Width="300"/>
        </StackPanel>
        <DataGrid Name="fileGrid" Grid.Row="2" AutoGenerateColumns="False" HeadersVisibility="Column" CanUserAddRows="False">
            <DataGrid.Columns>
                <DataGridCheckBoxColumn Header="Sélection" Binding="{Binding IsSelected, Mode=TwoWay}" Width="Auto"/>
                <DataGridTextColumn Header="Type" Binding="{Binding Type}" Width="Auto"/>
                <DataGridTextColumn Header="Nom actuel" Binding="{Binding OldName}" Width="*"/>
                <DataGridTextColumn Header="Nouveau nom" Binding="{Binding NewName}" Width="*"/>
            </DataGrid.Columns>
        </DataGrid>
        <TextBlock Grid.Row="3" Text="Générez les noms puis cliquez sur Renommer pour les fichiers sélectionnés." Margin="0,10,0,0"/>
    </Grid>
</Window>
"@

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Références aux éléments
$btnSelectFolder = $window.FindName("btnSelectFolder")
$btnRefresh = $window.FindName("btnRefresh")
$btnGenerateNames = $window.FindName("btnGenerateNames")
$btnRename = $window.FindName("btnRename")
$chkSelectAll = $window.FindName("chkSelectAll")
$fileGrid = $window.FindName("fileGrid")
$txtFilter = $window.FindName("txtFilter")

$global:currentFolder = $MediaScanPath
$global:fileList = @()
$global:fileData = @()
$global:filteredData = @()

$txtFilter.Add_TextChanged({ Set-Filter })

$chkSelectAll.Add_Checked({
    foreach ($item in $global:filteredData) { $item.IsSelected = $true }
    $fileGrid.ItemsSource = $null
    $fileGrid.ItemsSource = $global:filteredData
})

$chkSelectAll.Add_Unchecked({
    foreach ($item in $global:filteredData) { $item.IsSelected = $false }
    $fileGrid.ItemsSource = $null
    $fileGrid.ItemsSource = $global:filteredData
})

$btnSelectFolder.Add_Click({
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($folderDialog.ShowDialog() -eq "OK") {
        $global:currentFolder = $folderDialog.SelectedPath
        Get-MediaFiles
    }
})

$btnRefresh.Add_Click({
    Get-MediaFiles
})

$btnGenerateNames.Add_Click({
    foreach ($item in $global:filteredData) {
        if ($item.IsSelected) {
            $fileName = [System.IO.Path]::GetFileName($item.OldName)
            switch (Get-MediaType -fileName $item.OldName)
            {
                'TvShow' { $item.NewName = Get-TvShowEpisodeTitleFromMovieDB -fileName $fileName }
                'Movie' {
                    $SelMovie = Get-MovieItemFromMovieDB -fileName $fileName;
                    if ($SelMovie -eq $null) { }
                    elseif ($SelMovie.GetType().FullName -eq "System.String") { $item.NewName = $SelMovie }
                    elseif ($SelMovie.GetType().FullName -eq "System.Object[]") { $item.NewName = $SelMovie[$SelMovie.Count -1] }
                    }
                Default {}
            }
        }
    }
    $fileGrid.ItemsSource = $null
    $fileGrid.ItemsSource = $global:filteredData
})

$btnRename.Add_Click({
    foreach ($item in $global:filteredData) {
        if ($item.IsSelected -and $item.NewName -and (Test-Path $item.OldName)) {
            New-Item -ItemType Directory -Path (Split-Path $item.NewName -Parent) -Force | Out-Null
            Move-Item -Path $item.OldName -Destination $item.NewName -Force
        }
    }
    [System.Windows.MessageBox]::Show("Renommage terminé.")
    Get-MediaFiles
})

#endregion

# Chargement initial avec le répertoire par défaut
Get-MediaFiles

$window.ShowDialog()
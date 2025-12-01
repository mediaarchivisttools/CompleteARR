# 🎯 CompleteARR - Your Automated Media Librarian

CompleteARR is a powerful tool that automatically organizes your TV shows and movies in **Sonarr** and **Radarr**. Think of it as your personal media librarian that:

- **Sorts new content** into the right categories (Family, Adult, Anime)
- **Moves completed shows** to your main library when they're ready
- **Keeps your library clean** by hiding incomplete content
- **Helps protect younger viewers** by keeping age-appropriate content separate
- **Monitors all bonus content** (specials) after regular episodes are complete

## 🎪 How It Works

### 🎬 For TV Shows (Sonarr):
- **AutoSorter**: Automatically sorts new shows into the right "set" (Family, Adult, Anime) based on content ratings, genres, and tags
- **Series Engine**: Moves completed shows from "Incomplete" to "Complete" sets and monitors special episodes

### 🎥 For Movies (Radarr):
- **AutoSorter**: Sorts new movies from unsorted quality profiles into the right categories
- **Film Engine**: Ensures movies stay in their correct folders based on quality profile mappings

## 📁 Project Structure

### 🎬 Sonarr Scripts:
- **`CompleteARR_SONARR_Launcher.ps1`** - Runs all Sonarr tools
- **`CompleteARR_SONARR_AutoSorter.ps1`** - Sorts new shows into the right sets
- **`CompleteARR_SONARR_SeriesEngine.ps1`** - Manages show completion and special episode monitoring

### 🎥 Radarr Scripts:
- **`CompleteARR_RADARR_Launcher.ps1`** - Runs all Radarr tools
- **`CompleteARR_RADARR_AutoSorter.ps1`** - Sorts movies into the right categories
- **`CompleteARR_RADARR_FilmEngine.ps1`** - Enforces profile-to-folder mappings

### 🛠️ Helper Tools:
- **`CompleteARR_FetchInfo_Launcher.ps1`** - Essential setup tool that shows your current Quality Profiles and Root Folders
- **`CompleteARR_Launch_All_Scripts.ps1`** - Runs the full suite of CompleteARR tools

### ⚙️ Configuration:
- **`CompleteARR_SONARR_Settings.yml`** - Sonarr-specific configuration
- **`CompleteARR_RADARR_Settings.yml`** - Radarr-specific configuration

## 🚀 Getting Started

### Step 1: Prerequisites
- **Sonarr** (for TV shows) and/or **Radarr** (for movies) installed and running
- **PowerShell 7.0** or newer
- Your media library set up with quality profiles and root folders

### Step 2: Set Up Your Media System

#### 📺 For Sonarr (TV Shows):
Create these quality profile pairs (Incomplete/Complete):

- **Family**: `Incomplete - Family` / `Complete - Family`
- **Adult**: `Incomplete - Adult` / `Complete - Adult`  
- **Anime**: `Incomplete - Anime` / `Complete - Anime`
- **Anime Family**: `Incomplete - Anime Family` / `Complete - Anime Family`
- **Unsorted**: For new content (AutoSorter source)

#### 🎬 For Radarr (Movies):
Create these quality profiles:

- **Family**: `Family Default Settings`
- **Adult**: `Adult Default Settings`
- **Anime**: `Anime Default Settings`
- **Anime Family**: `Anime Family Default Settings`
- **Unsorted**: For new content (AutoSorter source)

### Step 3: Easy Setup with FetchInfo Tool

**Run `CompleteARR_FetchInfo_Launcher.ps1` first!**

This tool will:
- Connect to your Sonarr and Radarr instances
- Show you all your quality profiles and root folders
- Generate a log file with all the information you need for configuration

### Step 4: Configure Your Settings

#### For TV Shows (Sonarr):
1. Copy `CompleteARR_SONARR_Settings.example.yml` to `CompleteARR_SONARR_Settings.yml`
2. Edit the file with your information from the FetchInfo tool
3. Set up your Sonarr URL and API key
4. Configure your sets with PascalCase keys:
   ```yaml
   SortTargets:
     Adult:
       QualityProfile: "Incomplete - Adult"
       RootFolder: "/data/Media/Show Collection - Incomplete Default"
     Family:
       QualityProfile: "Incomplete - Family" 
       RootFolder: "/data/Media/Show Collection - Incomplete Family"
   ```

#### For Movies (Radarr):
1. Copy `CompleteARR_RADARR_Settings.example.yml` to `CompleteARR_RADARR_Settings.yml`
2. Edit the file with your information from the FetchInfo tool
3. Set up your Radarr URL and API key
4. Configure your sort targets with PascalCase keys

### Step 5: Run CompleteARR

**For Everything:**
- Run `CompleteARR_Launch_All_Scripts.ps1`

**For TV Shows Only:**
- Run `CompleteARR_SONARR_Launcher.ps1`

**For Movies Only:**
- Run `CompleteARR_RADARR_Launcher.ps1`

## 🔧 Key Features

### Smart Auto-Sorting
- Uses content ratings (TV-Y, PG, R, etc.)
- Checks genres and networks for family-friendly hints
- Respects manual overrides with tags like `LibraryOverride-Family`
- Extra careful with family content using strict mode

### Completion Tracking (Sonarr)
- Only shows complete content to your users
- Automatically monitors special episodes when shows are complete
- 15-day grace period for new missing episodes
- Clear separation between incomplete and complete libraries

### Profile Enforcement (Radarr)
- Ensures movies stay in correct folders based on quality profiles
- Maintains library organization automatically
- Prevents movies from drifting between categories

### Safety First
- Strict family mode for extra protection
- Clear separation between age groups
- Conservative handling of unrated content

## 🏷️ Manual Overrides with Tags

You can manually control sorting using these Sonarr/Radarr tags (create them in Settings > Tags):

- **`LibraryOverride-Family`** → Forces to Family target
- **`LibraryOverride-Adult`** → Forces to Adult target  
- **`LibraryOverride-Xplicit`** → Treats as explicit content
- **`LibraryOverride-Anime`** → Forces anime mode (still uses rating for family/adult decision)

## 💡 Tips for Success

1. **Use the FetchInfo tool first** - It makes setup much easier!
2. **Set up your Plex/Jellyfin libraries** to only include the "Complete" folders
3. **Use Sonarr/Radarr tags** to manually override sorting when needed
4. **Run CompleteARR regularly** (set up a scheduled task)
5. **Check the logs** in the `CompleteARR_Logs` folder if something doesn't work
6. **Start with dry runs** by setting `DryRun: true` in your settings

## 🔒 What CompleteARR Does NOT Do

- ❌ **Does NOT download content** (that's Sonarr/Radarr's job)
- ❌ **Does NOT search for torrents or NZBs**
- ❌ **Does NOT access the internet** except to talk to your Sonarr/Radarr
- ❌ **Does NOT modify your media files**

## ⚖️ Legal & Responsible Use

**Important:** CompleteARR is designed to help you organize media you are legally allowed to have. This includes:

- Media you purchased and ripped yourself
- Personal recordings where allowed by law
- Content you are explicitly licensed to download

**You are responsible for:**
- Ensuring your setup complies with local laws
- Respecting terms of service for any services you use
- Only using CompleteARR with content you have rights to

By using CompleteARR, you agree to use it responsibly and legally.

## 🆘 Need Help?

1. **Check the logs** in the `CompleteARR_Logs` folder
2. **Review your settings files** - make sure everything matches your Sonarr/Radarr setup
3. **Use the FetchInfo tool** to verify your configuration
4. **Start with dry runs** by setting `DryRun: true` in your settings
5. **Verify your quality profiles and root folders** match what's in your settings

---

**CompleteARR v1.0** 🚀


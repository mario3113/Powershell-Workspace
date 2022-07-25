Try {
	Import-Module PSGalleryModule -ErrorAction Stop
} Catch {
	Throw "Before using this script, please install PSGalleryModule: Install-Module PSGalleryModule"
}

# FMAutoTagPhotos
Automatically adds keyword tags to photos and videos in a collection.

## Background
Like many other FileMeta tools, FMAutoTagPhotos assumes that you have a master library of photos, videos, and audio clips. A common scenario is that a family member copies a selection of photos out of the master library onto a local drive, a thumb drive, optical disk, etc. for a specific use. Perhaps it's a birthday, wedding, or special event where you want a photo slide show. Perhaps you're sending photos to a friend or relative. Or perhaps you collected photos to print. Regardless, the collection remains valuable after the event but it's cluttering your storage to keep redundant copies of the photos around.

## Description
FMAutoTagPhotos is a command-line utility that takes a collection of photos in a folder or on a thumbdrive, locates the matching originals in the master library, and tags the originals with the specified keyword. Tags are stored as "keywords" in JPEG EXIF metadata.

The tool can also generate a list of all keywords that have been used in your collection.

FMAutoTagPhotos relies on Windows Search to locate the files. Therefore the master library must be included in the Windows Search index. It relies on EXIF metadata to ensure that the match is correct. This works well for photos that originate from digital cameras and phones. It does not work effectively with scanned images because they lack detailed metadata such as the date/time when the photo was taken.The collection of photos to tag need not be indexed.

## Use
FMAutoTagPhotos is a windows command-line program. Enter "FMAutoTagPhotos -h" for the syntax.

## Build
FMAutoTagPhotos is built using Microsoft Visual Studio Express 2013. No third party librararies or other components are required to build it.

## License
Copyright 2015 Brandt Redd

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

[http://www.apache.org/licenses/LICENSE-2.0](http://www.apache.org/licenses/LICENSE-2.0)

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.

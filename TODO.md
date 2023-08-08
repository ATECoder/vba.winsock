# TODO

## Tests

## Fixes

## Updates
Test:
* Add socket and TCP client
	* Add Prologix unit tests
		* before all: establish a 'connection'
		* inconclusive on all tests if not connected
		* read some Prologix properties.
	* Add ieee488 unit tests: *opc?; 
		* connect on before all
			* determine if Prologix connected
			* how to quickly determine if instrument is connected
		* inconclusive if not connected.
		* use read raw
		* use buffer read
		* use TCP client.
MVVM:
	* use MVVM View Model for the unit test sheet
	* use MVVM View Model for the identity sheet.
Upload release to GitHub:
	* add deploy and localize scripts.
	* [release a build artifact asset on git hub]
	* [gh release upload]
	* [gh release create]
	
[release a build artifact asset on git hub]: release https://stackoverflow.com/questions/5207269/how-to-release-a-build-artifact-asset-on-github-with-a-script
[gh release upload]: https://cli.github.com/manual/gh_release_upload
[gh release create]: https://cli.github.com/manual/gh_release_create
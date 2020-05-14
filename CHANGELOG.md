# Change Log

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](http://keepachangelog.com/)
and this project adheres to [Semantic Versioning](http://semver.org/).

## [*Unreleased]

### Changed

* Switched to `eslint` and removed `tslint`
* Switched to *renamed pacakge* `@fluentui/react-northstar`
* Breaking change, `inTeams()` now returns a `Promise<Boolean>`

### Deleted

* Removed unnecessary constructor
* Removed `fontSize` from state
* Removed `pageFontSize` protected method
* Removed the property interface
* Removed the `setValidityState` method

## [*2.0.0*] - <*2020-03-05*>

### Changed

* Replaced `msteams-ui-components-react` with `@fluentui/react`
* Updated dependecies (`@microsoft/teams-js`)
* Refactored packages into `peerDependencies`

## [*1.1.1*] - <*2019-04-17*>

### Fixes

* Fixed an issue with React typings version

## [*1.1.0*] - <*2019-03-24*>

### Changed

* Updated Typescript, React, teams-js and msteams-ui-components-react libraries

## [*1.0.0*] - <*2018-08-13*>

### Changed

* Updated version of @microsoft/teams-js

## [*0.0.4*] - <*2018-07-30*>

### Changed

* Updated code to reflect linting changes

### Fixed

* Fixed missing import to the @microsoft/teams-js library

### Added

* Travis-ci integration
* Linting added (npm run lint)

## [*0.0.3*] - <*2018-07-28*>

### Changed

* Changed signature for getQueryVariable and updateTheme

## [*0.0.2*] - <*2018-07-28*>

### Added

* More inline documentation

### Changed

* Included comments in output
* Updated node packages
* Updated method signatures
* Updated tsconfig.json with stricter checking

## [*0.0.1*] - <*2018-07-28*>

### Added

* Initial release

# XLSBuild
Extension for MSBuild providing Tasks for managing XLSX files with VBA modules

## Overview and Purpose
MSBuild is a full-featured build system like GNU make or ant, designed around and for the .NET framework. MSBuild makes automating even complex build and deployment processes very straightforward.

Microsoft Excel is a widely used spreadsheet application that has a powerful built-in macro system called Visual Basic for Applications or VBA. Many companies have built extensive and complex business critical applications in Excel using VBA.

VBA wasn't originally intended for these purposes, so using VBA with modern best-practice techniques like source control, continuous integration and test-driven development is very difficult.

This library exposes to MSBuild a number of useful tasks for managing VBA modules inside Excel spreadsheets. The intent is for developers to be able to extract their VBA from Excel to enable it to be unit-tested and checked into source control, merging the VBA back into the Excel templates as part of the build process.

## Quick Start

## Task Reference


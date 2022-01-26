# VBA API Proof Of Concept
A proof of concept of VBA being used to read API (fully commented with useful information)

In this I extract the last 100 records for MSFT shares taken at 5 minute intervals. The share, interval and number of records can be easily altered in the VBA code (use the API documentation at https://rapidapi.com/ for the available parameters).

This tool uses the following; 1 - The Jason.Converter.bas module from the VBA-JSON project on GitHHub. This can be found at; https://github.com/VBA-tools/VBA-JSON 2 - A reference to the "Microsoft WinHTTP Services, version 5.1" library 3 - A reference to the "Microsoft Scripting Runtime" library 4 - An account (free) at https://rapidapi.com/ 5 - Subscribe to the Alpha Vantage API at RapidAPI (Free, but necessary step)

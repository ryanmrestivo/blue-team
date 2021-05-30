# Find-LOLBAS

A simple Powershell script for enumerating living off the land binaries and scripts on a system.

## Why

Manually verifying if the binaries or scripts are on the system <br>
would take a while, with automating the process this increases overall productivity
of redteamers <br>
who need to quickly bypass applocker or need to execute code in unique ways.

## How to Use?

By simply running the script the rest is taken care of! <br>
The output will be on the screen for you to assess, it will be in the format <br>
of the Binary or Script name, path, and an example command utilizing it.

## License

This project is licensed under the BSD 3-Clause License -
see the [License](LICENSE) file for details

## Acknowledgments

This project wouldn't be possible without the [LOLBAS](https://github.com/LOLBAS-Project/LOLBAS) project.

### Roadmap

- [ ] Add option to run script by executing C# code in Powershell
- [ ] Add option to allow user to encode payload by loading Crypt32.dll

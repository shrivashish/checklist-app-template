# Checklist

## Deployment Steps

### Prerequisites
- Node: https://nodejs.org/en/download/
- Git.

### Steps
1. Clone this repo.
```
git clone <checklist_repo>.git
```
2. Change directory
```
cd Checklist
```
3. Install dependencies
```
npm install
```
4. Change `<packageId>` in actionManifest.json
5. Deploy the package 
```
npm run create
```
6. `<packageId>.zip` will be generated in output folder. Upload the zip in teams.

## Scripts

### ```npm run build```
Build the app and generate output folder.

### ```npm run start```
Build the app and generate output folder along with map files for all JS. Also watch the input files and rebuild if there is any change.

### ```npm run zip```
Zip the content of output folder and create file `ActionPackage.zip`.

### ```npm run create```
Upload the `ActionPackage.zip` to ActionPlatfrom and generate `<packageId>.zip` file in output folder.

### ```npm run update```
Update the ActionPackage to ActionPlatfrom.

### ```npm run inner-loop```
Replace `<packageId>` with actual package id mentioned in action manifest in package.json before run this command. This command is useful for devlopment as the package is serve from output folder instead of action service.

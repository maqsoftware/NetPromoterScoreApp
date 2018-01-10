(function () {
    "use strict";
    WinJS.Namespace.define('VaultManager', {    
    asyncVaultLoad: function () {
        return new WinJS.Promise(function (complete, error, progress) {
            var vault = new Windows.Security.Credentials.PasswordVault();
            // any call to the password vault will load the vault 
            var creds = vault.retrieveAll();
            complete();
        });
    },
    addCredential: function (AWSAccessKey, AWSSecretKey) {
        try {            
            var resource = "AWSKeys";            
            if (resource === "" || AWSAccessKey === "" || AWSSecretKey === "") {
                console.log("Resouce, AccessKey and SecretKey are required when adding a credential");
                VaultManager.deleteCredential();
                return;
            }
            var vault = new Windows.Security.Credentials.PasswordVault();
            var cred = new Windows.Security.Credentials.PasswordCredential(resource, AWSAccessKey, AWSSecretKey);
            vault.add(cred);
            console.log("Credential saved successfully.");
        }
        catch (e) {
            console.log(e.message + " " + e.toString());
            return;
        }
    },
    readCredential: function () {
        try {
            var oCreds = {
                "resource": "AWSKeys",
                "AWSAccessKey": "",
                "AWSSecretKey": "",
            };
            var resource = "AWSKeys";           
            var vault = new Windows.Security.Credentials.PasswordVault();            
            var creds = null;
            if (vault !== null) {
                creds = vault.findAllByResource(resource);
            }
            if (creds !== null) {
                creds = vault.retrieve(resource, creds[0].userName);
                oCreds.AWSAccessKey = creds.userName;                
                oCreds.AWSSecretKey = creds.password;
                console.log("Read credential succeeds");
            }            
            else {
                console.log("Credential not found.");
            }
        }
        catch (e) {
            console.log(e.message);
            return oCreds;
        }
        return oCreds;

    },
    deleteCredential: function () {
        try {
            var resource = "AWSKeys", creds = "", AccessKey = "";
            var vault = new Windows.Security.Credentials.PasswordVault();
            if (vault !== null) {                
                creds = vault.findAllByResource(resource);
                if (creds !== null) {
                    creds = vault.retrieve(resource, creds[0].userName);
                    vault.remove(creds);
                    console.log("Credential removed successfully.");
                }                
            }
        } catch (e) {
            console.log(e.message);
            return;
        }

    }
    });    
})();
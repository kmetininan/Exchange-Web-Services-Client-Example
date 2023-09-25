/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.metininan.ewsattachget;

import com.microsoft.aad.msal4j.ClientCredentialFactory;
import com.microsoft.aad.msal4j.ClientCredentialParameters;
import com.microsoft.aad.msal4j.ConfidentialClientApplication;
import com.microsoft.aad.msal4j.IAuthenticationResult;
import java.io.IOException;
import java.net.URI;
import java.net.URISyntaxException;
import java.util.Collections;
import java.util.Properties;
import java.util.concurrent.CompletableFuture;
import java.util.logging.Level;
import java.util.logging.Logger;
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.misc.ConnectingIdType;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceLocalException;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.misc.ImpersonatedUserId;
import microsoft.exchange.webservices.data.property.complex.FolderId;
import microsoft.exchange.webservices.data.property.complex.Mailbox;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;

/**
 *
 * @author METIN
 */
public class NewClass {

    private static String authority;
    private static String clientId;
    private static String secret;
    private static String scope;
    private static final String EWS_URL = "<exchange_server_ul>/EWS/Exchange.asmx";
 
    public static void main(String[] args) {
        
        
        
        try {
            setUpSampleData();
        } catch (IOException ex) {
            Logger.getLogger(NewClass.class.getName()).log(Level.SEVERE, null, ex);
        }

       
       try (ExchangeService service = getAuthenticatedService("canias@exc.canias.com")) {
            try {
                listInboxMessages(service, "Canias01@exc.canias.com");
            } catch (Exception ex) {
                Logger.getLogger(NewClass.class.getName()).log(Level.SEVERE, null, ex);
            }
    }   catch (Exception ex) {
            Logger.getLogger(NewClass.class.getName()).log(Level.SEVERE, null, ex);
        }
       
    }

    public static ExchangeService getAuthenticatedService(String token, String senderAddr) throws URISyntaxException, Exception {
        ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);

        service.getHttpHeaders().put("Authorization", "Bearer " + token);
        service.getHttpHeaders().put("X-AnchorMailbox", senderAddr);
        //service.setWebProxy(new WebProxy(proxyHost, proxyPort));
        service.setUrl(new URI(EWS_URL));
        service.setImpersonatedUserId(new ImpersonatedUserId(ConnectingIdType.PrincipalName, senderAddr));
        return service;
    }
    
    public static ExchangeService getAuthenticatedService(String senderAddr) throws URISyntaxException, Exception {
        ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
        
        service.setCredentials(new WebCredentials("<username>", "<password>","exc"));
           
            service.setTraceEnabled(true);
          
        
        service.setUrl(new URI(EWS_URL));
        //service.setImpersonatedUserId(new ImpersonatedUserId(ConnectingIdType.PrincipalName, senderAddr));
        return service;
    }

    private static IAuthenticationResult getAccessTokenByClientCredentialGrant() throws Exception {

        ConfidentialClientApplication app = ConfidentialClientApplication.builder(
                clientId,
                ClientCredentialFactory.createFromSecret(secret))
                .authority(authority)
                .build();

        // With client credentials flows the scope is ALWAYS of the shape "resource/.default", as the
        // application permissions need to be set statically (in the portal), and then granted by a tenant administrator
        ClientCredentialParameters clientCredentialParam = ClientCredentialParameters.builder(
                Collections.singleton(scope))
                .build();

        CompletableFuture<IAuthenticationResult> future = app.acquireToken(clientCredentialParam);
        return future.get();
    }

    private static void setUpSampleData() throws IOException {
        // Load properties file and set properties used throughout the sample
        Properties properties = new Properties();
        properties.load(Thread.currentThread().getContextClassLoader().getResourceAsStream("application.properties"));
        authority = properties.getProperty("AUTHORITY");
        clientId = properties.getProperty("CLIENT_ID");
        secret = properties.getProperty("SECRET");
        scope = properties.getProperty("SCOPE");
    }

    public static void listInboxMessages(ExchangeService service, String mailboxAddr) throws Exception {
        ItemView view = new ItemView(50);
        Mailbox mb = new Mailbox(mailboxAddr);
        FolderId folder = new FolderId(WellKnownFolderName.Inbox, mb);
        FindItemsResults<Item> result = service.findItems(folder, view);
        result.forEach(i -> {
            try {
                System.out.println("subject=" + i.getSubject());
            } catch (ServiceLocalException e) {
                e.printStackTrace();
            }
        });
    }
}

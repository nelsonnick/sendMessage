package com.wts.util;

import me.chanjar.weixin.cp.api.WxCpInMemoryConfigStorage;

public class WxConfig extends WxCpInMemoryConfigStorage {
    private volatile static WxConfig wxConfig;

    private WxConfig() {
        this.setCorpId(ParamesAPI.CorpID);
        this.setCorpSecret(ParamesAPI.Secret);
        this.setAgentId(ParamesAPI.AgentId);
//        this.setToken(ParamesAPI.token);
//        this.setAesKey(ParamesAPI.encodingAESKey);
    }

    public static WxConfig getMe() {
        if (wxConfig == null) {
            synchronized (WxConfig.class) {
                if (wxConfig == null) {
                    wxConfig = new WxConfig();
                }
            }
        }
        return wxConfig;
    }


}
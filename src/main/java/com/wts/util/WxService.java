package com.wts.util;

import me.chanjar.weixin.cp.api.WxCpServiceImpl;

public class WxService extends WxCpServiceImpl {
  private volatile static WxService wxService;

  private WxService() {
    this.setWxCpConfigStorage(WxConfig.getMe());
  }

  public static WxService getMe() {
    if (wxService == null) {
      synchronized (WxService.class) {
        if (wxService == null) {
          wxService = new WxService();
        }
      }
    }
    return wxService;
  }


}
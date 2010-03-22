package org.no.ip.bca.exchange

import com.microsoft.schemas.exchange.services._2006.types.{ExchangeVersionType, RequestServerVersionDocument}

private[exchange] trait VersionHelper {
  protected def requestVersion = {
    val requestServerVersion = RequestServerVersionDocument.Factory.newInstance
    requestServerVersion setRequestServerVersion {
      val requestServerVersion = RequestServerVersionDocument.RequestServerVersion.Factory.newInstance
      requestServerVersion setVersion ExchangeVersionType.EXCHANGE_2007
      requestServerVersion
    }
    requestServerVersion
  }
}
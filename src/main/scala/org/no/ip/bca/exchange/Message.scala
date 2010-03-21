package org.no.ip.bca.exchange

import java.util.Calendar
import com.microsoft.schemas.exchange.services._2006.messages._
import com.microsoft.schemas.exchange.services._2006.types._
import org.no.ip.bca.exchange.web.services.ExchangeServicesStub

class Message private[exchange](ews: ExchangeServicesStub, itemId: ItemIdType) {
  private lazy val item = {
    val getItemDocument = GetItemDocument.Factory.newInstance
    getItemDocument setGetItem {
      val getItemType = GetItemType.Factory.newInstance
      val itemIds = getItemType.addNewItemIds
      itemIds.addNewItemId
      itemIds.setItemIdArray(0, itemId.copy.asInstanceOf[ItemIdType])
      getItemType.addNewItemShape setBaseShape DefaultShapeNamesType.ALL_PROPERTIES
      getItemType
    }
    val ret = ews.getItem(getItemDocument, null, null, null, requestVersion)
    ret.getGetItemResponse.getResponseMessages.getGetItemResponseMessageArray(0).getItems.getMessageArray(0)
  }
  
  def getId = itemId.getId
  def getSubject = item.getSubject
  def getSize = item.getSize
  def getDateTimeCreated = item.getDateTimeCreated.clone.asInstanceOf[Calendar]
  def getDateTimeSent = item.getDateTimeSent.clone.asInstanceOf[Calendar]
  def isRead = item.getIsRead
  def getSensitivity = item.getSensitivity
  def getBody = item.getBody
  
  def getDateTimeReceived = item.getDateTimeReceived.clone.asInstanceOf[Calendar]
  
  private lazy val headers = Map(item.getInternetMessageHeaders.getInternetMessageHeaderArray map { header =>
    header.getHeaderName -> header.getStringValue
  }: _*)
  def getHeaders = headers
  def getConversationIndex = Helper.asString(item.getConversationIndex)
  def getFromEmail = item.getFrom.getMailbox.getEmailAddress
  def getFromName = item.getFrom.getMailbox.getName
  def getMessageType = item
  
  private def requestVersion = {
    val requestServerVersion = RequestServerVersionDocument.Factory.newInstance
    requestServerVersion setRequestServerVersion {
      val requestServerVersion = RequestServerVersionDocument.RequestServerVersion.Factory.newInstance
      requestServerVersion setVersion ExchangeVersionType.EXCHANGE_2007
      requestServerVersion
    }
    requestServerVersion
  }
  
  override def toString = getSubject + " (" + getSize + ") >> " + getConversationIndex + "\n" + item
}
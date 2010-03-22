package org.no.ip.bca.exchange

import java.util.Calendar
import com.microsoft.schemas.exchange.services._2006.messages._
import com.microsoft.schemas.exchange.services._2006.types._
import org.no.ip.bca.exchange.web.services.ExchangeServicesStub

class Message private[exchange](ews: ExchangeServicesStub, itemId: ItemIdType) extends VersionHelper {
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
  def getDateTimeReceived = item.getDateTimeReceived.clone.asInstanceOf[Calendar]
  def getDateTimeSent = item.getDateTimeSent.clone.asInstanceOf[Calendar]
  def getToRecipients = {
    Helper.option(item.getToRecipients) map {
      _.getMailboxArray
    } getOrElse {new Array(0)} map { mailbox =>
      val name = Helper.option(mailbox.getName)
      val email = mailbox.getEmailAddress
      Recipient(email, name)
    }
  }
  def isDraft = item.getIsDraft
  def isRead = item.getIsRead
  def getSensitivity = item.getSensitivity
  def getBody = item.getBody
  
  private lazy val headers = Map(item.getInternetMessageHeaders.getInternetMessageHeaderArray map { header =>
    header.getHeaderName -> header.getStringValue
  }: _*)
  def getHeaders = headers
  def getConversationIndex = Helper.asString(item.getConversationIndex)
  def getFrom = {
    val mailbox = item.getFrom.getMailbox
    val email = mailbox.getEmailAddress
    val name = Helper.option(mailbox.getName)
    Recipient(email, name)
  }
  def getMessageType = item
  
  def hasAttachments = {
    val attachments = item.getAttachments
    attachments != null && (!attachments.getFileAttachmentArray.isEmpty || !attachments.getItemAttachmentArray.isEmpty)
  }
  def getItemAttachments = {
    Helper.option(item.getAttachments) map { a =>
      a.getItemAttachmentArray
    } getOrElse new Array(0)
  }
  def getFileAttachments = {
    Helper.option(item.getAttachments) map { a =>
      a.getFileAttachmentArray map { a => new FileAttachment(ews, a.getAttachmentId) }
    } getOrElse new Array(0)
  }
  
  override def toString = getSubject + " (" + getSize + ") >> " + getConversationIndex + "\n" + item
}
package org.no.ip.bca.exchange

import com.microsoft.schemas.exchange.services._2006.messages._
import com.microsoft.schemas.exchange.services._2006.types._
import org.no.ip.bca.exchange.web.services.ExchangeServicesStub

class FileAttachment private[exchange](ews: ExchangeServicesStub, attachmentId: AttachmentIdType) extends VersionHelper {
  private lazy val item = {
    val doc = GetAttachmentDocument.Factory.newInstance
    doc setGetAttachment {
      val getAtt = GetAttachmentType.Factory.newInstance
      val attachmentIds = getAtt.addNewAttachmentIds
      attachmentIds.addNewAttachmentId
      attachmentIds.setAttachmentIdArray(0, attachmentId.copy.asInstanceOf[AttachmentIdType])
      
      val shape = getAtt.addNewAttachmentShape
      shape setIncludeMimeContent true
      val pth = PathToUnindexedFieldType.Factory.newInstance
      pth setFieldURI UnindexedFieldURIType.ITEM_MIME_CONTENT
      shape.addNewAdditionalProperties setPathArray Array(pth)
      getAtt
    }
    val ret = ews.getAttachment(doc, null, null, null, requestVersion)
    ret.getGetAttachmentResponse.getResponseMessages.getGetAttachmentResponseMessageArray(0).getAttachments.getFileAttachmentArray(0)
  }
  
  def getContent = item.getContent
  def getContentId = item.getContentId
  def getContentType = item.getContentType
  def getName = item.getName
}
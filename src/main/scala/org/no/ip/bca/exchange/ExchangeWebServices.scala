package org.no.ip.bca.exchange;

import java.nio.charset.Charset
import java.util.concurrent.TimeUnit
import com.microsoft.schemas.exchange.services._2006.messages._
import com.microsoft.schemas.exchange.services._2006.types._
import org.apache.axis2.transport.http.{HTTPConstants, HttpTransportProperties}
import org.apache.commons.codec.binary.Base64
import org.apache.commons.httpclient.util.EncodingUtil
import org.no.ip.bca.exchange.web.services.ExchangeServicesStub

private [exchange] object Helper {
  val UTF_8 = Charset.forName("UTF-8")
  def asString(bytes: Array[Byte]) = new String(Base64.encodeBase64(bytes), UTF_8)
  def asBytes(string: String) = Base64.decodeBase64(string getBytes UTF_8)
}

class ExchangeWebServices(server: String, domain: String, username: String, password: String) {
  private val ews = new ExchangeServicesStub("https://" + server + "/EWS/Exchange.asmx")
  ews._getServiceClient.getOptions.setProperty(HTTPConstants.AUTHENTICATE, {
    val auth = new HttpTransportProperties.Authenticator
    auth setHost server
    auth setDomain domain
    auth setUsername username
    auth setPassword password
    auth
  })
  setTimeout(5, TimeUnit.MINUTES)
  
  def setTimeout(time: Long, unit: TimeUnit) = {
    ews._getServiceClient.getOptions.setTimeOutInMilliSeconds(unit.toMillis(time))
  }
  
  def getInbox = {
    val folderId = NonEmptyArrayOfBaseFolderIdsType.Factory.newInstance
    folderId.addNewDistinguishedFolderId setId DistinguishedFolderIdNameType.INBOX
    val ret = ews.getFolder(getFolderReq(folderId), null, null, null, requestVersion)
    val folder = ret.getGetFolderResponse.getResponseMessages.getGetFolderResponseMessageArray(0).getFolders.getFolderArray(0)
    new Folder(ews, folder.getFolderId)
  }
  
  def getFolderById(id: String) = {
    val folderId = NonEmptyArrayOfBaseFolderIdsType.Factory.newInstance
    folderId.addNewFolderId setId id
    val ret = ews.getFolder(getFolderReq(folderId), null, null, null, requestVersion)
    val folder = ret.getGetFolderResponse.getResponseMessages.getGetFolderResponseMessageArray(0).getFolders.getFolderArray(0)
    new Folder(ews, folder.getFolderId)
  }
  
  def getItemAttachmentById(id: String) = {
    getAttachmentById(id).getItemAttachmentArray(0)
  }
  
  def getFileAttachmentById(id: String) = {
    getAttachmentById(id).getFileAttachmentArray(0)
  }
  
  private def getAttachmentById(id: String) = {
    val doc = GetAttachmentDocument.Factory.newInstance
    val getAtt = doc.addNewGetAttachment
    getAtt.addNewAttachmentIds.addNewAttachmentId setId id
    val shape = getAtt.addNewAttachmentShape
    shape setIncludeMimeContent true
    val pth = PathToUnindexedFieldType.Factory.newInstance
    pth setFieldURI UnindexedFieldURIType.ITEM_SIZE
    shape.addNewAdditionalProperties setPathArray Array(pth)
    val ret = ews.getAttachment(doc, null, null, null, requestVersion)
    ret.getGetAttachmentResponse.getResponseMessages.getGetAttachmentResponseMessageArray(0).getAttachments
  }
  
  private def getFolderReq(folderId: NonEmptyArrayOfBaseFolderIdsType) = {
    val getFolderReq = GetFolderDocument.Factory.newInstance
    getFolderReq setGetFolder {
      val getFolderType = GetFolderType.Factory.newInstance
      getFolderType setFolderShape {
        val folderResponseShapeType = FolderResponseShapeType.Factory.newInstance
        folderResponseShapeType setBaseShape DefaultShapeNamesType.ID_ONLY
        folderResponseShapeType
      }
      getFolderType setFolderIds folderId
      getFolderType
    }
    getFolderReq
  }
  
  private def requestVersion = {
    val requestServerVersion = RequestServerVersionDocument.Factory.newInstance
    requestServerVersion setRequestServerVersion {
      val requestServerVersion = RequestServerVersionDocument.RequestServerVersion.Factory.newInstance
      requestServerVersion setVersion ExchangeVersionType.EXCHANGE_2007
      requestServerVersion
    }
    requestServerVersion
  }
}

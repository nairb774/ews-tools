package org.no.ip.bca.exchange

import com.microsoft.schemas.exchange.services._2006.messages._
import com.microsoft.schemas.exchange.services._2006.types._
import org.no.ip.bca.exchange.web.services.ExchangeServicesStub

class Folder private[exchange](ews: ExchangeServicesStub, folderId: FolderIdType) extends VersionHelper {
  private lazy val folder = {
    val getFolderDocument = GetFolderDocument.Factory.newInstance
    getFolderDocument setGetFolder {
      val getFolderType = GetFolderType.Factory.newInstance
      getFolderType setFolderIds {
        val nonEmptyArrayOfBaseFolderIdsType = NonEmptyArrayOfBaseFolderIdsType.Factory.newInstance
        nonEmptyArrayOfBaseFolderIdsType.addNewFolderId
        nonEmptyArrayOfBaseFolderIdsType.setFolderIdArray(0, folderId.copy.asInstanceOf[FolderIdType])
        nonEmptyArrayOfBaseFolderIdsType
      }
      getFolderType.addNewFolderShape setBaseShape DefaultShapeNamesType.DEFAULT
      getFolderType
    }
    val ret = ews.getFolder(getFolderDocument, null, null, null, requestVersion)
    ret.getGetFolderResponse.getResponseMessages.getGetFolderResponseMessageArray(0).getFolders.getFolderArray(0)
  }
  def getId = folderId.getId
  def getDisplayName = folder.getDisplayName
  def getChildFolderCount = folder.getChildFolderCount
  def getParentFolder = {
    val parentId = folder.getParentFolderId
    if (parentId == null)
      None
    else
      Some(new Folder(ews, parentId))
  }
  def getTotalCount = folder.getTotalCount
  def getUnreadCount = folder.getUnreadCount
  def getFolders = {
    val findFolder = FindFolderDocument.Factory.newInstance
    findFolder setFindFolder {
      val findFolderType = FindFolderType.Factory.newInstance
      findFolderType setTraversal FolderQueryTraversalType.SHALLOW
      findFolderType.addNewFolderShape setBaseShape DefaultShapeNamesType.ID_ONLY
      findFolderType setParentFolderIds nonEmptyArrayOfBaseFolderIdsType
      findFolderType
    }
    val ret = ews.findFolder(findFolder, null, null, null, requestVersion)
    val folders = ret.getFindFolderResponse.getResponseMessages.getFindFolderResponseMessageArray(0).getRootFolder.getFolders.getFolderArray
    if (folders == null) {
      new Array[Folder](0)
    } else {
      folders map { folder => new Folder(ews, folder.getFolderId) }
    }
  }
  
  def getItems(restriction: Option[Expression]) = {
    val findItem = FindItemDocument.Factory.newInstance
    findItem setFindItem {
      val findItemType = FindItemType.Factory.newInstance
      findItemType setTraversal ItemQueryTraversalType.SHALLOW
      findItemType.addNewItemShape setBaseShape DefaultShapeNamesType.ID_ONLY
      restriction foreach { restriction => findItemType setRestriction AsRestriction(restriction) }
      findItemType setParentFolderIds nonEmptyArrayOfBaseFolderIdsType
      findItemType
    }
    val ret = ews.findItem(findItem, null, null, null, requestVersion)
    val items = ret.getFindItemResponse.getResponseMessages.getFindItemResponseMessageArray(0).getRootFolder.getItems.getMessageArray
    items map { item => new Message(ews, item.getItemId) }
  }
  
  private def nonEmptyArrayOfBaseFolderIdsType = {
    val nonEmptyArrayOfBaseFolderIdsType = NonEmptyArrayOfBaseFolderIdsType.Factory.newInstance
    nonEmptyArrayOfBaseFolderIdsType.addNewFolderId // Expand array so we can add clone:
    nonEmptyArrayOfBaseFolderIdsType.setFolderIdArray(0, folder.getFolderId.copy.asInstanceOf[FolderIdType])
    nonEmptyArrayOfBaseFolderIdsType
  }
  
  override def toString = getDisplayName
}
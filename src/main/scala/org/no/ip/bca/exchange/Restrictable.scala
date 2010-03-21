package org.no.ip.bca.exchange

import com.microsoft.schemas.exchange.services._2006.types._

private[exchange] object RestrictionsHelper {
  def ISO_DATE_FORMAT = new java.text.SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss.SSSZ")
  
  private val altTypes = Map(
    "PathToUnindexedFieldType" -> "FieldURI"
  )
  
  /** This metod helps with fixing: http://social.technet.microsoft.com/Forums/en/exchangesvrdevelopment/thread/a02dea87-7b64-44f4-92f4-d5dff3fcf55a */
  def fix[O <: org.apache.xmlbeans.XmlObject](obj: O): O ={
    val typ = obj.schemaType
    val typName = typ.getName
    val origLocalPart = typName.getLocalPart
    if (!origLocalPart.endsWith("Type")) throw new IllegalArgumentException(typName.toString)
    val localPart = altTypes.getOrElse(origLocalPart, origLocalPart.substring(0, origLocalPart.length - 4))
    obj.substitute(new javax.xml.namespace.QName(typName.getNamespaceURI, localPart), typ).asInstanceOf[O]
  }
}

abstract sealed trait DateField {
  import RestrictionsHelper._
  def path: PathToUnindexedFieldType
  private[DateField] def pushPath[T <: TwoOperandExpressionType](expr: T): T = {
    expr setPath path
    expr setPath fix(expr.getPath)
    expr
  }
  
  def >=(date: java.util.Date) = {
    val expression = pushPath(IsGreaterThanOrEqualToType.Factory.newInstance)
    expression.addNewFieldURIOrConstant.addNewConstant setValue ISO_DATE_FORMAT.format(date)
    new GreaterEqualThan(expression)
  }
  def >(date: java.util.Date) = {
    val expression = pushPath(IsGreaterThanType.Factory.newInstance)
    expression.addNewFieldURIOrConstant.addNewConstant setValue ISO_DATE_FORMAT.format(date)
    new GreaterThan(expression)
  }
  def <=(date: java.util.Date) = {
    val expression = pushPath(IsLessThanOrEqualToType.Factory.newInstance)
    expression.addNewFieldURIOrConstant.addNewConstant setValue ISO_DATE_FORMAT.format(date)
    new LessEqualThan(expression)
  }
  def <(date: java.util.Date) = {
    val expression = pushPath(IsLessThanType.Factory.newInstance)
    expression.addNewFieldURIOrConstant.addNewConstant setValue ISO_DATE_FORMAT.format(date)
    new LessThan(expression)
  }
  def ~(date: java.util.Date) = {
    val expression = pushPath(IsEqualToType.Factory.newInstance)
    expression.addNewFieldURIOrConstant.addNewConstant setValue ISO_DATE_FORMAT.format(date)
    new EqualTo(expression)
  }
  def !=(date: java.util.Date) = {
    val expression = pushPath(IsNotEqualToType.Factory.newInstance)
    expression.addNewFieldURIOrConstant.addNewConstant setValue ISO_DATE_FORMAT.format(date)
    new NotEqualTo(expression)
  }
}
final object ItemDateTimeCreated extends DateField {
  def path = {
    val path = PathToUnindexedFieldType.Factory.newInstance
    path setFieldURI UnindexedFieldURIType.ITEM_DATE_TIME_CREATED
    path
  }
}
final object ItemDateTimeRecieved extends DateField {
  def path = {
    val path = PathToUnindexedFieldType.Factory.newInstance
    path setFieldURI UnindexedFieldURIType.ITEM_DATE_TIME_RECEIVED
    path
  }
}
final object ItemDateTimeSent extends DateField {
  def path = {
    val path = PathToUnindexedFieldType.Factory.newInstance
    path setFieldURI UnindexedFieldURIType.ITEM_DATE_TIME_SENT
    path
  }
}
sealed abstract trait BooleanField {
  import RestrictionsHelper._
  def path: PathToUnindexedFieldType
  private[BooleanField] def pushPath[T <: TwoOperandExpressionType](expr: T): T = {
    expr setPath path
    expr setPath fix(expr.getPath)
    expr
  }
  def ~(bool: Boolean) = {
    val expression = pushPath(IsEqualToType.Factory.newInstance)
    expression.addNewFieldURIOrConstant.addNewConstant setValue bool.toString
    new EqualTo(expression)
  }
  def !=(bool: Boolean) = {
    val expression = pushPath(IsNotEqualToType.Factory.newInstance)
    expression.addNewFieldURIOrConstant.addNewConstant setValue bool.toString
    new NotEqualTo(expression)
  }
}
final object MessageIsRead extends BooleanField {
  def path = {
    val path = PathToUnindexedFieldType.Factory.newInstance
    path setFieldURI UnindexedFieldURIType.MESSAGE_IS_READ
    path
  }
}
private[exchange] final object AsRestriction {
  import RestrictionsHelper._
  def apply(expr: Expression) = {
    val restrictionType = RestrictionType.Factory.newInstance
    restrictionType setSearchExpression expr.expr
    restrictionType setSearchExpression fix(restrictionType.getSearchExpression)
    restrictionType
  }
}
abstract sealed class Expression(private[exchange] val expr: SearchExpressionType) {
  import RestrictionsHelper._
  
  def &&(other: Expression) = {
    val exprs = Array(this, other)
    val andType = AndType.Factory.newInstance
    andType.setSearchExpressionArray(exprs.map(_.expr))
    for(val i <- 0 until exprs.length) {
      andType.setSearchExpressionArray(i, fix(andType.getSearchExpressionArray(i)))
    }
    new And(andType)
  }
  def ||(other: Expression) = {
    val exprs = Array(this, other)
    val orType = OrType.Factory.newInstance
    orType.setSearchExpressionArray(exprs.map(_.expr))
    for(val i <- 0 until exprs.length) {
      orType.setSearchExpressionArray(i, fix(orType.getSearchExpressionArray(i)))
    }
    new Or(orType)
  }
}
final class And private[exchange](expr: AndType) extends Expression(expr)
final class Or private[exchange](expr: OrType) extends Expression(expr)
final class GreaterEqualThan private[exchange](expr: IsGreaterThanOrEqualToType) extends Expression(expr)
final class GreaterThan private[exchange](expr: IsGreaterThanType) extends Expression(expr)
final class LessEqualThan private[exchange](expr: IsLessThanOrEqualToType) extends Expression(expr)
final class LessThan private[exchange](expr: IsLessThanType) extends Expression(expr)
final class EqualTo private[exchange](expr: IsEqualToType) extends Expression(expr)
final class NotEqualTo private[exchange](expr: IsNotEqualToType) extends Expression(expr)

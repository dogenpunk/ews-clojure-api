(ns exchange.ews.contact
  (:require [exchange.ews.authentication :refer [service-instance]]
            [exchange.ews.util :refer [enum-id-cond]])
  (:import (microsoft.exchange.webservices.data.core.service.item Contact)
           (microsoft.exchange.webservices.data.core.service.schema ContactSchema)
           (microsoft.exchange.webservices.data.property.complex ItemId)
           (microsoft.exchange.webservices.data.core.enumeration.property WellKnownFolderName)
           (microsoft.exchange.webservices.data.search ItemView)
           (microsoft.exchange.webservices.data.search.filter SearchFilter$IsEqualTo)
           (microsoft.exchange.webservices.data.core.enumeration.service ConflictResolutionMode)
           (microsoft.exchange.webservices.data.property.complex CompleteName
                                                                 EmailAddress
                                                                 EmailAddressCollection
                                                                 EmailAddressDictionary
                                                                 ImAddressDictionary
                                                                 ItemAttachment
                                                                 ItemId
                                                                 PhoneNumberDictionary
                                                                 StringList)))

(defmulti get-contact
  type)

(defmethod get-contact ItemId
  [id]
  (Contact/bind @service-instance id))

(defmethod get-contact java.lang.String
  [id]
  (Contact/bind @service-instance (ItemId/getItemIdFromString id)))

(defn get-contact-by-name
  [name]
  (let [search-filter (SearchFilter$IsEqualTo. ContactSchema/DisplayName name)
        items (.findItems @service-instance WellKnownFolderName/Contacts search-filter (ItemView. 1))]
    (first (.getItems items))))

(defn set-contact-display-name
  [id display-name]
  (let [contact (get-contact id)]
    (.setDisplayName contact display-name)
    (.update contact ConflictResolutionMode/AutoResolve)))

(defn set-contact-given-name
  [id given-name]
  (let [contact (get-contact id)]
    (.setGivenName contact given-name)
    (.update contact ConflictResolutionMode/AutoResolve)))

(defn set-contact-surname
  [id surname]
  (let [contact (get-contact id)]
    (.setSurname contact surname)
    (.update contact ConflictResolutionMode/AutoResolve)))

(ns exchange.ews.contacts
  (:require [exchange.ews.authentication :refer [server-instance]])
  (:import (microsoft.exchange.webservices.data.core.service.item Contact)
           (microsoft.exchange.webservices.data.core.enumeration.service ConflictResolutionMode)
           (microsoft.exchange.webservices.data.property.complex CompleteName
                                                                 EmailAddress
                                                                 EmailAddressCollection
                                                                 EmailAddressDictionary
                                                                 ImAddressDictionary
                                                                 ItemAttachment
                                                                 ItemId
                                                                 PhoneNumberDictionary
                                                                 PhysicalAddress
                                                                 StringList)))

(defn set-contact-given-name
  [id given-name]
  (let [contact (Contact/bind @service-instance (ItemId/getItemIdFromString id))]
    (.setGivenName contact display-name)
    (.update contact ConflictResolutionMode/AutoResolve)))

(defn set-email-address
  [id email-address]
  (let [contact (Contact/bind @service-instance (ItemId/getItemIdFromString id))]
    (.set)))

{-
Under BSD 3-Clause license, (c) 2009 Doug Beardsley <mightybyte@gmail.com>, (c) 2009-2012 Stephen Blackheath <http://blacksapphire.com/antispam/>, (c) 2009 Gregory Collins, (c) 2008 Evan Martin <martine@danga.com>, (c) 2009 Matthew Pocock <matthew.pocock@ncl.ac.uk>, (c) 2007-2009 Galois Inc., (c) 2010 Kevin Jardine, (c) 2012 Simon Hengel

From https://hackage.haskell.org/package/hexpat-0.20.13
     https://github.com/the-real-blackh/hexpat/blob/master/Text/XML/Expat/SAX.hs#L227
-}
module Codec.Xlsx.Parser.Stream.HexpatInternal (parseBuf) where

import Control.Monad
import Text.XML.Expat.SAX
import qualified Data.ByteString.Internal as I
import Data.Bits
import Data.Int
import Data.ByteString.Internal (c_strlen)
import Data.Word
import Foreign.C
import Foreign.ForeignPtr
import Foreign.Ptr
import Foreign.Storable

{-# SCC parseBuf #-}
parseBuf :: (GenericXMLString tag, GenericXMLString text) =>
            ForeignPtr Word8 -> CInt -> (Ptr Word8 -> Int -> IO (a, Int)) -> IO [(SAXEvent tag text, a)]
parseBuf buf _ processExtra = withForeignPtr buf $ \pBuf -> doit [] pBuf 0
  where
    roundUp32 offset = (offset + 3) .&. complement 3
    doit acc pBuf offset0 = offset0 `seq` do
        typ <- peek (pBuf `plusPtr` offset0 :: Ptr Word32)
        (a, offset) <- processExtra pBuf (offset0 + 4)
        case typ of
            0 -> return (reverse acc)
            1 -> do
                nAtts <- peek (pBuf `plusPtr` offset :: Ptr Word32)
                let pName = pBuf `plusPtr` (offset + 4)
                lName <- fromIntegral <$> c_strlen pName
                let name = gxFromByteString $ I.fromForeignPtr buf (offset + 4) lName
                (atts, offset') <- foldM (\(atts, offset) _ -> do
                        let pAtt = pBuf `plusPtr` offset
                        lAtt <- fromIntegral <$> c_strlen pAtt
                        let att = gxFromByteString $ I.fromForeignPtr buf offset lAtt
                            offset' = offset + lAtt + 1
                            pValue = pBuf `plusPtr` offset'
                        lValue <- fromIntegral <$> c_strlen pValue
                        let value = gxFromByteString $ I.fromForeignPtr buf offset' lValue
                        return ((att, value):atts, offset' + lValue + 1)
                    ) ([], offset + 4 + lName + 1) [1,3..nAtts]
                doit ((StartElement name (reverse atts), a) : acc) pBuf (roundUp32 offset')
            2 -> do
                let pName = pBuf `plusPtr` offset
                lName <- fromIntegral <$> c_strlen pName
                let name = gxFromByteString $ I.fromForeignPtr buf offset lName
                    offset' = offset + lName + 1
                doit ((EndElement name, a) : acc) pBuf (roundUp32 offset')
            3 -> do
                len <- fromIntegral <$> peek (pBuf `plusPtr` offset :: Ptr Word32)
                let text = gxFromByteString $ I.fromForeignPtr buf (offset + 4) len
                    offset' = offset + 4 + len
                doit ((CharacterData text, a) : acc) pBuf (roundUp32 offset')
            4 -> do
                let pEnc = pBuf `plusPtr` offset
                lEnc <- fromIntegral <$> c_strlen pEnc
                let enc = gxFromByteString $ I.fromForeignPtr buf offset lEnc
                    offset' = offset + lEnc + 1
                    pVer = pBuf `plusPtr` offset'
                pVerFirst <- peek (castPtr pVer :: Ptr Word8)
                (mVer, offset'') <- case pVerFirst of
                    0 -> return (Nothing, offset' + 1)
                    1 -> do
                        lVer <- fromIntegral <$> c_strlen (pVer `plusPtr` 1)
                        return (Just $ gxFromByteString $ I.fromForeignPtr buf (offset' + 1) lVer, offset' + 1 + lVer + 1)
                    _ -> error "hexpat: bad data from C land"
                cSta <- peek (pBuf `plusPtr` offset'' :: Ptr Int8)
                let sta = if cSta < 0  then Nothing else
                          if cSta == 0 then Just False else
                                            Just True
                doit ((XMLDeclaration enc mVer sta, a) : acc) pBuf (roundUp32 (offset'' + 1))
            5 -> doit ((StartCData, a) : acc) pBuf offset
            6 -> doit ((EndCData, a) : acc) pBuf offset
            7 -> do
                let pTarget = pBuf `plusPtr` offset
                lTarget <- fromIntegral <$> c_strlen pTarget
                let target = gxFromByteString $ I.fromForeignPtr buf offset lTarget
                    offset' = offset + lTarget + 1
                    pData = pBuf `plusPtr` offset'
                lData <- fromIntegral <$> c_strlen pData
                let dat = gxFromByteString $ I.fromForeignPtr buf offset' lData
                doit ((ProcessingInstruction target dat, a) : acc) pBuf (roundUp32 (offset' + lData + 1))
            8 -> do
                let pText = pBuf `plusPtr` offset
                lText <- fromIntegral <$> c_strlen pText
                let text = gxFromByteString $ I.fromForeignPtr buf offset lText
                doit ((Comment text, a) : acc) pBuf (roundUp32 (offset + lText + 1))
            _ -> error "hexpat: bad data from C land"

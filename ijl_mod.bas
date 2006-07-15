Attribute VB_Name = "IJL"
Rem               INTEL CORPORATION PROPRIETARY INFORMATION
Rem  This software is supplied under the terms of a license agreement or
Rem  nondisclosure agreement with Intel Corporation and may not be copied
Rem  or disclosed except in accordance with the terms of that agreement.
Rem      Copyright (c) 1998 Intel Corporation. All Rights Reserved.
Rem
Rem  File:      ijl.bas
Rem  Purpose:   Intel(R) JPEG Library Visual Basic interface module
Rem  Version:   1.51

Option Explicit

Global Const JBUFSIZE  As Long = 4096

Rem  Purpose:     Stores library version info.
Rem  Example:
Rem   major           - 1
Rem   minor           - 0
Rem   build           - 1
Rem   Name            - "ijl10.dll"
Rem   Version         - "1.0.1 Beta 1"
Rem   InternalVersion - "1.0.1.1"
Rem   BuildDate       - "Sep 22 1998"
Rem   CallConv        - "DLL"
Public Type IJLibVersion
  major           As Long
  minor           As Long
  build           As Long
  Name            As Long 'pointer to C-style string
  Version         As Long 'pointer to C-style string
  InternalVersion As Long 'pointer to C-style string
  BuildDate       As Long 'pointer to C-style string
  CallConv        As Long 'pointer to C-style string
End Type


Rem Purpose:     Keep coordinates for rectangle region of image
Rem Context:     Used to specify roi
Public Type IJL_RECT
  Left   As Long
  Top    As Long
  Right  As Long
  Bottom As Long
End Type


Rem  Purpose:    Possible types of data read/write/other operations to be
Rem              performed by the functions IJL_Read and IJL_Write.
Rem              See the Developer's Guide for details on appropriate usage.
Rem    IJL_JFILE_XXXXXXX   Indicates JPEG data in a stdio file.
Rem    IJL_JBUFF_XXXXXXX   Indicates JPEG data in an addressable buffer.
Public Enum IJLIOTYPE
  IJL_SETUP = -1
  ' Read JPEG parameters (i.e., height, width, channels, sampling, etc.)
  ' from a JPEG bit stream.
  IJL_JFILE_READPARAMS = 0
  IJL_JBUFF_READPARAMS = 1
  ' Read a JPEG Interchange Format image.
  IJL_JFILE_READWHOLEIMAGE = 2
  IJL_JBUFF_READWHOLEIMAGE = 3
  ' Read JPEG tables from a JPEG Abbreviated Format bit stream.
  IJL_JFILE_READHEADER = 4
  IJL_JBUFF_READHEADER = 5
  ' Read image info from a JPEG Abbreviated Format bit stream.
  IJL_JFILE_READENTROPY = 6
  IJL_JBUFF_READENTROPY = 7
  ' Write an entire JFIF bit stream.
  IJL_JFILE_WRITEWHOLEIMAGE = 8
  IJL_JBUFF_WRITEWHOLEIMAGE = 9
  ' Write a JPEG Abbreviated Format bit stream.
  IJL_JFILE_WRITEHEADER = 10
  IJL_JBUFF_WRITEHEADER = 11
  ' Write image info to a JPEG Abbreviated Format bit stream.
  IJL_JFILE_WRITEENTROPY = 12
  IJL_JBUFF_WRITEENTROPY = 13
  ' Scaled Decoding Options:
  ' Reads a JPEG image scaled to 1/2 size.
  IJL_JFILE_READONEHALF = 14
  IJL_JBUFF_READONEHALF = 15
  ' Reads a JPEG image scaled to 1/4 size.
  IJL_JFILE_READONEQUARTER = 16
  IJL_JBUFF_READONEQUARTER = 17
  ' Reads a JPEG image scaled to 1/8 size.
  IJL_JFILE_READONEEIGHTH = 18
  IJL_JBUFF_READONEEIGHTH = 19
  ' Reads an embedded thumbnail from a JFIF bit stream.
  IJL_JFILE_READTHUMBNAIL = 20
  IJL_JBUFF_READTHUMBNAIL = 21
End Enum


Rem  Purpose:     Possible color space formats.
Rem  Note these formats do *not* necessarily denote
Rem  the number of channels in the color space.
Rem  There exists separate "channel" fields in the
Rem  JPEG_CORE_PROPERTIES data structure specifically
Rem  for indicating the number of channels in the
Rem  JPEG and/or DIB color spaces.
Public Enum IJL_COLOR
  IJL_RGB = 1        ' Red-Green-Blue color space.
  IJL_BGR = 2        ' Reversed channel ordering from IJL_RGB.
  IJL_YCBCR = 3      ' Luminance-Chrominance color space as defined by CCIR Recommendation 601.
  IJL_G = 4          ' Grayscale color space.
  IJL_RGBA_FPX = 5   ' FlashPix RGB 4 channel color space that has pre-multiplied opacity.
  IJL_YCBCRA_FPX = 6 ' FlashPix YCbCr 4 channel color space that has pre-multiplied opacity.
  IJL_OTHER = 255    ' Some other color space not defined by the IJL.
End Enum


Rem  Purpose:     Possible subsampling formats used in the JPEG.
Public Enum IJL_JPGSUBSAMPLING
  IJL_NONE = 0  ' Corresponds to "No Subsampling". Valid on a JPEG w/ any number of channels.
  IJL_411 = 1   ' Valid on a JPEG w/ 3 channels.
  IJL_422 = 2   ' Valid on a JPEG w/ 3 channels.
  IJL_4114 = 3  ' Valid on a JPEG w/ 4 channels.
  IJL_4224 = 4  ' Valid on a JPEG w/ 4 channels.
End Enum


Rem  Purpose:     Possible subsampling formats used in the DIB.
Public Enum IJL_DIBSUBSAMPLING
  IJL_SNONE = 0  ' Corresponds to "No Subsampling". Valid on a DIB w/ any number of channels.
  IJL_S422 = 2   ' Valid on a DIB with YCbYCr color.
End Enum


Rem  Purpose:     Stores Huffman table information in a fast-to-use format.
Rem  Context:     Used by Huffman encoder/decoder to access Huffman table
Rem               data.  Raw Huffman tables are formatted to fit this
Rem               structure prior to use.
Rem  Fields:
Rem    huff_class  0 == DC Huffman or lossless table, 1 == AC table.
Rem    ident       Huffman table identifier, 0-3 valid (Extended Baseline).
Rem    huffelem    Huffman elements for codes <= 8 bits long;
Rem                contains both zero run-length and symbol length in bits.
Rem    huffval     Huffman values for codes 9-16 bits in length.
Rem    mincode     Smallest Huffman code of length n.
Rem    maxcode     Largest Huffman code of length n.
Rem    valptr      Starting index into huffval[] for symbols of length k.
Public Type HUFFMAN_TABLE
  huff_class         As Long
  ident              As Long
  huffelem(0 To 255) As Long
  huffval(0 To 255)  As Integer
  mincode(0 To 16)   As Integer
  maxcode(0 To 17)   As Integer
  valptr(0 To 16)    As Integer
End Type


Rem  Purpose:     Stores pointers to JPEG-binary spec compliant
Rem               Huffman table information.
Rem  Context:     Used by interface and table methods to specify encoder
Rem               tables to generate and store JPEG images.
Rem  Fields:
Rem    bits        Points to number of codes of length i (<=16 supported).
Rem    vals        Value associated with each Huffman code.
Rem    hclass      0 == DC table, 1 == AC table.
Rem    ident       Specifies the identifier for this table; 0-3 for extended JPEG compliance.
Public Type JPEGHuffTable
  bits   As Long
  vals   As Long
  hclass As Byte
  ident  As Byte
  ' IJL use 8 byte pack structures
  pad0   As Byte
  pad1   As Byte
End Type


Rem  Purpose:     Stores quantization table information in a
Rem               fast-to-use format.
Rem  Context:     Used by quantizer/dequantizer to store formatted
Rem               quantization tables.
Rem  Fields:
Rem    precision   0 => elements contains 8-bit elements,
Rem                1 => elements contains 16-bit elements.
Rem    ident       Table identifier (0-3).
Rem    elements    Pointer to 64 table elements + 16 extra elements to catch
Rem                input data errors that may cause malfunction of the
Rem                Huffman decoder.
Rem    elarray     Space for elements (see above) plus 8 bytes to align
Rem                to a quadword boundary.
Public Type QUANT_TABLE
  precision        As Long
  ident            As Long
  Elements         As Long
  elarray(0 To 83) As Integer
End Type


Rem  Purpose:     Stores pointers to JPEG binary spec compliant
Rem               quantization table information.
Rem  Context:     Used by interface and table methods to specify encoder
Rem               tables to generate and store JPEG images.
Rem  Fields:
Rem    quantizer   Zig-zag order elements specifying quantization factors.
Rem    ident       Specifies identifier for this table.
Rem                0-3 valid for Extended Baseline JPEG compliance.
Public Type JPEGQuantTable
  quantizer As Long
  ident     As Byte
  ' IJL use 8 byte pack structures
  pad0      As Byte
  pad1      As Byte
  pad2      As Byte
End Type


Rem  Purpose:     One frame-component structure is allocated per component
Rem               in a frame.
Rem
Rem  Context:     Used by Huffman decoder to manage components.
Rem
Rem  Fields:
Rem    ident       Component identifier.  The tables use this ident to
Rem                determine the correct table for each component.
Rem    hsampling   Horizontal subsampling factor for this component,
Rem                1-4 are legal.
Rem    vsampling   Vertical subsampling factor for this component,
Rem                1-4 are legal.
Rem    quant_sel   Quantization table selector.  The quantization table
Rem                used by this component is determined via this selector.
Public Type FRAME_COMPONENT
  ident     As Long
  hsampling As Long
  vsampling As Long
  quant_sel As Long
End Type


Rem  Purpose:     Stores frame-specific data.
Rem
Rem  Context:     One Frame structure per image.
Rem
Rem  Fields:
Rem    precision       Sample precision in bits.
Rem    width           Width of the source image in pixels.
Rem    height          Height of the source image in pixels.
Rem    MCUheight       Height of a frame MCU.
Rem    MCUwidth        Width of a frame MCU.
Rem    max_hsampling   Max horiz sampling ratio of any component in the frame.
Rem    max_vsampling   Max vert sampling ratio of any component in the frame.
Rem    ncomps          Number of components/channels in the frame.
Rem    horMCU          Number of horizontal MCUs in the frame.
Rem    totalMCU        Total number of MCUs in the frame.
Rem    comps           Array of 'ncomps' component descriptors.
Rem    restart_interv  Indicates number of MCUs after which to restart the entropy parameters.
Rem    SeenAllDCScans  Used when decoding Multiscan images to determine if all channels of an image have been decoded.
Rem    SeenAllACScans  (See SeenAllDCScans)
Public Type ijl_FRAME
  precision      As Long
  width          As Long
  Height         As Long
  MCUheight      As Long
  MCUwidth       As Long
  max_hsampling  As Long
  max_vsampling  As Long
  ncomps         As Long
  horMCU         As Long
  totalMCU       As Long
  comps          As Long
  restart_interv As Long
  SeenAllDCScans As Long
  SeenAllACScans As Long
End Type


Rem  Purpose:     One scan-component structure is allocated per component
Rem               of each scan in a frame.
Rem
Rem  Context:     Used by Huffman decoder to manage components within scans.
Rem
Rem  Fields:
Rem    comp        Component number, index to the comps member of FRAME.
Rem    hsampling   Horizontal sampling factor.
Rem    vsampling   Vertical sampling factor.
Rem    dc_table    DC Huffman table pointer for this scan.
Rem    ac_table    AC Huffman table pointer for this scan.
Rem    quant_table Quantization table pointer for this scan.
Public Type SCAN_COMPONENT
  comp       As Long
  hsampling  As Long
  vsampling  As Long
  dc_table   As Long
  ac_table   As Long
  quantTable As Long
End Type


Rem  Purpose:     One SCAN structure is allocated per scan in a frame.
Rem  Fields:
Rem    ncomps          Number of image components in a scan, 1-4 legal.
Rem    gray_scale      If TRUE, decode only the Y channel.
Rem    start_spec      Start coefficient of spectral or predictor selector.
Rem    end_spec        End coefficient of spectral selector.
Rem    approx_high     High bit position in successive approximation
Rem                    Progressive coding.
Rem    approx_low      Low bit position in successive approximation
Rem                    Progressive coding.
Rem    restart_interv  Restart interval, 0 if disabled.
Rem    curxMCU         Next horizontal MCU index to be processed after
Rem                    an interrupted SCAN.
Rem    curyMCU         Next vertical MCU index to be processed after
Rem                    an interrupted SCAN.
Rem    dc_diff         Array of DC predictor values for DPCM modes.
Rem    comps           Array of ncomps SCAN_COMPONENT component identifiers.
Public Type SCAN
  ncomps          As Long
  gray_scale      As Long
  start_spec      As Long
  end_spec        As Long
  approx_high     As Long
  approx_low      As Long
  restart_interv  As Long
  curxMCU         As Long
  curyMCU         As Long
  dc_diff(0 To 3) As Long
  comps           As Long
End Type


Rem  Purpose:     Possible algorithms to be used to perform the discrete
Rem               cosine transform (DCT).
Rem
Rem  Fields:
Rem    IJL_AAN     The AAN (Arai, Agui, and Nakajima) algorithm from
Rem                Trans. IEICE, vol. E 71(11), 1095-1097, Nov. 1988.
Public Enum DCTTYPE
  IJL_AAN = 0
  IJL_IPP = 1
End Enum


Rem Purpose:            -  Possible algorithms to be used to perform upsampling
Rem Fields:
Rem  IJL_BOX_FILTER      - the algorithm is simple replication of the input pixel
Rem                        onto the corresponding output pixels (box filter);
Rem  IJL_TRIANGLE_FILTER - 3/4 * nearer pixel + 1/4 * further pixel in each
Rem                        dimension
Public Enum upsampling_type
  IJL_BOX_FILTER = 0
  IJL_TRIANGLE_FILTER = 1
End Enum


Rem Purpose:     Stores current conditions of sampling. Only for upsampling
Rem              with triangle filter is used now.
Rem
Rem Fields:
Rem  top_row        - pointer to buffer with MCUs, that are located above than
Rem                   current row of MCUs;
Rem  cur_row        - pointer to buffer with current row of MCUs;
Rem  bottom_row     - pointer to buffer with MCUs, that are located below than
Rem                   current row of MCUs;
Rem  last_row       - pointer to bottom boundary of last row of MCUs
Rem  cur_row_number - number of row of MCUs, that is decoding;
Rem  user_interrupt - field to store jprops->interrupt, because of we prohibit
Rem                   interrupts while top row of MCUs is upsampling.
Public Type SAMPLING_STATE
  top_row        As Long
  cur_row        As Long
  bottom_row     As Long
  last_row       As Long
  cur_row_number As Long
End Type


Rem  Purpose:     Possible types of processors.
Rem               Note that the enums are defined in ascending order
Rem               depending upon their various IA32 instruction support.
Rem    IJL_OTHER_PROC
Rem      Does not support the CPUID instruction and assumes no Pentium(R) processor instructions.
Rem
Rem    IJL_PENTIUM_PROC
Rem      Corresponds to an Intel(R) Pentium(R) processor(or a 100% compatible) that supports the
Rem      Pentium(R) processor instructions.
Rem
Rem    IJL_PENTIUM_PRO_PROC
Rem      Corresponds to an Intel(R) Pentium(R) Pro processor(or a 100% compatible) that supports the
Rem      Pentium(R) Pro processor instructions.
Rem
Rem    IJL_PENTIUM_PROC_MMX_TECH
Rem      Corresponds to an Intel(R) Pentium(R) processor with MMX(TM) technology (or a 100% compatible)
Rem      that supports the MMX(TM) instructions.
Rem
Rem    IJL_PENTIUM_II_PROC
Rem      Corresponds to an Intel(R) Pentium(R) II procesor  (or a 100% compatible) that supports both the
Rem      Pentium(R) Pro processor instructions and the MMX(TM) instructions.
Rem
Rem    IJL_PENTIUM_III_PROC
Rem      Corresponds to an Intel(R) Pentium(R) III processor
Rem
Rem    IJL_PENTIUM_4_PROC
Rem      Corresponds to an Intel(R) Pentium(R) 4 processor
Public Enum PROCESSOR_TYPE
  IJL_OTHER_PROC = 0
  IJL_PENTIUM_PROC = 1
  IJL_PENTIUM_PRO_PROC = 2
  IJL_PENTIUM_PROC_MMX_TECH = 3
  IJL_PENTIUM_II_PROC = 4
  IJL_PENTIUM_III_PROC = 5
  IJL_PENTIUM_4_PROC = 6
End Enum


Rem Purpose:     Stores data types: raw dct coefficients or raw sampled data.
Rem              Pointer to structure in JPEG_PROPERTIES is NULL, if any raw
Rem              data isn't request (DIBBytes!=NULL).
Rem Fields:
Rem  short* raw_ptrs[4] - pointers to buffers with raw data; one pointer corresponds one JPG component;
Rem  data_type          - 0 - raw dct coefficients, 1 - raw sampled data.
Public Type RAW_DATA_TYPES_STATE
  data_type As Long
  raw_ptrs(0 To 3)  As Long
End Type

Rem  Purpose:     Stores the decoder state information necessary to "jump"
Rem               to a particular MCU row in a compressed entropy stream.
Rem  Fields:
Rem    offset              Offset (in bytes) into the entropy stream from the beginning.
Rem    dcval1              DC val at the beginning of the MCU row for component 1.
Rem    dcval2              DC val at the beginning of the MCU row for component 2.
Rem    dcval3              DC val at the beginning of the MCU row for component 3.
Rem    dcval4              DC val at the beginning of the MCU row for component 4.
Rem    bit_buffer_64       64-bit Huffman bit buffer.  Stores current
Rem                        bit buffer at the start of a MCU row.
Rem                        Also used as a 32-bit buffer on 32-bit
Rem                        architectures.
Rem    bitbuf_bits_valid   Number of valid bits in the above bit buffer.
Rem    unread_marker       Have any markers been decoded but not
Rem                        processed at the beginning of a MCU row?
Rem                        This entry holds the unprocessed marker, or
Rem                        0 if none.
Public Type ENTROPYSTRUCT
  offset               As Long
  dcval1               As Long
  dcval2               As Long
  dcval3               As Long
  dcval4               As Long
  ' IJL use 8 byte pack structures
  pad0                 As Byte
  pad1                 As Byte
  pad2                 As Byte
  pad3                 As Byte
  bit_buffer_64        As Long
  bit_buffer_64_part_2 As Long
  bitbuf_bits_valid    As Long
  unread_marker        As Byte
  ' IJL use 8 byte pack structures
  pad4                 As Byte
  pad5                 As Byte
  pad6                 As Byte
End Type


Rem  Purpose:     Stores the active state of the IJL.
Rem  Fields:
Rem    bit_buffer_64           64-bit bitbuffer utilized by Huffman
Rem                            encoder/decoder algorithms utilizing routines
Rem                            designed for MMX(TM) technology.
Rem    bit_buffer_32           32-bit bitbuffer for all other Huffman encoder/decoder algorithms.
Rem    bitbuf_bits_valid       Number of bits in the above two fields that are valid.
Rem
Rem    cur_entropy_ptr         Current position (absolute address) in the entropy buffer.
Rem    start_entropy_ptr       Starting position (absolute address) of the entropy buffer.
Rem    end_entropy_ptr         Ending position (absolute address) of the entropy buffer.
Rem    entropy_bytes_processed Number of bytes actually processed(passed over) in the entropy buffer.
Rem    entropy_buf_maxsize     Max size of the entropy buffer.
Rem    entropy_bytes_left      Number of bytes left in the entropy buffer.
Rem    Prog_EndOfBlock_Run     Progressive block run counter.
Rem
Rem    DIB_ptr                 Temporary offset into the input/output DIB.
Rem
Rem    unread_marker           If a marker has been read but not processed, stick it in this field.
Rem    processor_type          (0, 1, or 2) == current processor does not
Rem                            support MMX(TM) instructions.
Rem                           (3 or 4) == current processor does
Rem                            support MMX(TM) instructions.
Rem    cur_scan_comp           On which component of the scan are we working?
Rem    file                    Process file handle, or 0x00000000 if no file is defined.
Rem    JPGBuffer               Entropy buffer (~4K).
Public Type STATE
  bit_buffer_64                As Long
  bit_buffer_64_part_2         As Long
  bit_buffer_32                As Long
  bitbuf_bits_valid            As Long
  cur_entropy_ptr              As Long
  start_entropy_ptr            As Long
  end_entropy_ptr              As Long
  entropy_bytes_processed      As Long
  entropy_buf_maxsize          As Long
  entropy_bytes_left           As Long
  Prog_EndOfBlock_Run          As Long
  DIB_ptr                      As Long
  unread_marker                As Byte
  ' IJL use 8 byte pack structures
  pad0                         As Byte
  pad1                         As Byte
  pad2                         As Byte
  proc_type                    As PROCESSOR_TYPE
  cur_scan_comp                As Long
  File                         As Long
  JPGBuffer(0 To JBUFSIZE - 1) As Byte
End Type


Rem  Purpose:     Advanced Control Option.  Do NOT modify.
Rem               WARNING:  Used for internal reference only.
Public Enum FAST_MCU_PROCESSING_TYPE
  IJL_NO_CC_OR_US = 0

  IJL_111_YCBCR_111_RGB = 1
  IJL_111_YCBCR_111_BGR = 2

  IJL_411_YCBCR_111_RGB = 3
  IJL_411_YCBCR_111_BGR = 4

  IJL_422_YCBCR_111_RGB = 5
  IJL_422_YCBCR_111_BGR = 6

  IJL_111_YCBCR_1111_RGBA_FPX = 7
  IJL_411_YCBCR_1111_RGBA_FPX = 8
  IJL_422_YCBCR_1111_RGBA_FPX = 9

  IJL_1111_YCBCRA_FPX_1111_RGBA_FPX = 10
  IJL_4114_YCBCRA_FPX_1111_RGBA_FPX = 11
  IJL_4224_YCBCRA_FPX_1111_RGBA_FPX = 12

  IJL_111_RGB_1111_RGBA_FPX = 13

  IJL_1111_RGBA_FPX_1111_RGBA_FPX = 14

  IJL_111_OTHER_111_OTHER = 15
  IJL_411_OTHER_111_OTHER = 16
  IJL_422_OTHER_111_OTHER = 17

  IJL_YCBYCR_YCBCR = 18   'encoding from YCbCr 422 format
  IJL_YCBCR_YCBYCR = 19   'decoding to YCbCr 422 format
End Enum


Rem  Purpose:     Stores low-level and control information.  It is used by
Rem               both the encoder and decoder.  An advanced external user
Rem               may access this structure to expand the interface
Rem               capability.
Rem  Fields:
Rem    iotype              IN:     Specifies type of data operation
Rem                                (read/write/other) to be
Rem                                performed by IJL_Read or IJL_Write.
Rem    roi                 IN:     Rectangle-Of-Interest to read from, or
Rem                                write to, in pixels.
Rem    dcttype             IN:     DCT alogrithm to be used.
Rem    fast_processing     OUT:    Supported fast pre/post-processing path. This is set by the IJL.
Rem    interrupt           IN:     Signals an interrupt has been requested.
Rem
Rem    DIBBytes            IN:     Pointer to buffer of uncompressed data.
Rem    DIBWidth            IN:     Width of uncompressed data.
Rem    DIBHeight           IN:     Height of uncompressed data.
Rem    DIBPadBytes         IN:     Padding (in bytes) at end of each row in the uncompressed data.
Rem    DIBChannels         IN:     Number of components in the uncompressed data.
Rem    DIBColor            IN:     Color space of uncompressed data.
Rem    DIBSubsampling      IN:     Required to be IJL_NONE or IJL_422.
Rem    DIBLineBytes        OUT:    Number of bytes in an output DIB line including padding.
Rem
Rem    JPGFile             IN:     Pointer to file based JPEG.
Rem    JPGBytes            IN:     Pointer to buffer based JPEG.
Rem    JPGSizeBytes        IN:     Max buffer size. Used with JPGBytes.
Rem                      OUT:      Number of compressed bytes written.
Rem    JPGWidth            IN:     Width of JPEG image.
Rem                      OUT:      After reading (except READHEADER).
Rem    JPGHeight           IN:     Height of JPEG image.
Rem                      OUT:      After reading (except READHEADER).
Rem    JPGChannels         IN:     Number of components in JPEG image.
Rem                      OUT:      After reading (except READHEADER).
Rem    JPGColor            IN:     Color space of JPEG image.
Rem    JPGSubsampling      IN:     Subsampling of JPEG image.
Rem                       OUT:     After reading (except READHEADER).
Rem    JPGThumbWidth       OUT:    JFIF embedded thumbnail width [0-255].
Rem    JPGThumbHeight      OUT:    JFIF embedded thumbnail height [0-255].
Rem
Rem    cconversion_reqd    OUT:    If color conversion done on decode, TRUE.
Rem    upsampling_reqd     OUT:    If upsampling done on decode, TRUE.
Rem    jquality            IN:     [0-100] where highest quality is 100.
Rem    jinterleaveType     IN/OUT: 0 => MCU interleaved file, and 1 => 1 scan per component.
Rem    numxMCUs            OUT:    Number of MCUs in the x direction.
Rem    numyMCUs            OUT:    Number of MCUs in the y direction.
Rem
Rem    nqtables            IN/OUT: Number of quantization tables.
Rem    maxquantindex       IN/OUT: Maximum index of quantization tables.
Rem    nhuffActables       IN/OUT: Number of AC Huffman tables.
Rem    nhuffDctables       IN/OUT: Number of DC Huffman tables.
Rem    maxhuffindex        IN/OUT: Maximum index of Huffman tables.
Rem    jFmtQuant           IN/OUT: Formatted quantization table info.
Rem    jFmtAcHuffman       IN/OUT: Formatted AC Huffman table info.
Rem    jFmtDcHuffman       IN/OUT: Formatted DC Huffman table info.
Rem
Rem    jEncFmtQuant        IN/OUT: Pointer to one of the above, or to externally persisted table.
Rem    jEncFmtAcHuffman    IN/OUT: Pointer to one of the above, or to externally persisted table.
Rem    jEncFmtDcHuffman    IN/OUT: Pointer to one of the above, or to externally persisted table.
Rem
Rem    use_external_qtables IN:    Set to default quantization tables. Clear to supply your own.
Rem    use_external_htables IN:    Set to default Huffman tables. Clear to supply your own.
Rem    rawquanttables      IN:     Up to 4 sets of quantization tables.
Rem    rawhufftables       IN:     Alternating pairs (DC/AC) of up to 4 sets of raw Huffman tables.
Rem    HuffIdentifierAC    IN:     Indicates what channel the user-supplied Huffman AC tables apply to.
Rem    HuffIdentifierDC    IN:     Indicates what channel the user-supplied Huffman DC tables apply to.
Rem
Rem    jframe              OUT:    Structure with frame-specific info.
Rem    needframe           OUT:    TRUE when a frame has been detected.
Rem
Rem    jscan               Persistence for current scan pointer when interrupted.
Rem
Rem    state               OUT:    Contains info on the state of the IJL.
Rem    SawAdobeMarker      OUT:    Decoder saw an APP14 marker somewhere.
Rem    AdobeXform          OUT:    If SawAdobeMarker TRUE, this indicates the JPEG color space given by that marker.
Rem
Rem    rowoffsets          Persistence for the decoder MCU row origins
Rem                        when decoding by ROI.  Offsets (in bytes
Rem                        from the beginning of the entropy data)
Rem                        to the start of each of the decoded rows.
Rem                        Fill the offsets with -1 if they have not
Rem                        been initalized and NULL could be the
Rem                        offset to the first row.
Rem
Rem    MCUBuf              OUT:    Quadword aligned internal buffer.
Rem                                Big enough for the largest MCU
Rem                                (10 blocks) with extra room for
Rem                                additional operations.
Rem    tMCUBuf             OUT:    Version of above, without alignment.
Rem
Rem    processor_type      OUT:    Determines type of processor found
Rem                                during initialization.
Rem
Rem    raw_coefs          IN/OUT   if !NULL, then pointer to vector of pointers
Rem                                (size = JPGChannels) to buffers for raw (short)
Rem                                dct coefficients. 1 pointer corresponds to one
Rem                                component;
Rem
Rem    progressive_found   OUT:    1 when progressive image detected.
Rem    coef_buffer         IN:     Pointer to a larger buffer containing
Rem                                frequency coefficients when they
Rem                                cannot be decoded dynamically
Rem                                (i.e., as in progressive decoding).
Rem
Rem    upsampling_type     IN:     Type of sampling:
Rem                                IJL_BOX_FILTER or IJL_TRIANGLE_FILTER.
Rem    SAMPLING_STATE*     OUT:    pointer to structure, describing current
Rem                                condition of upsampling
Rem
Rem    AdobeVersion       OUT      version field, if Adobe APP14 marker detected
Rem    AdobeFlags0        OUT      flags0 field, if Adobe APP14 marker detected
Rem    AdobeFlags1        OUT      flags1 field, if Adobe APP14 marker detected
Rem
Rem    jfif_app0_detected OUT:     1 - if JFIF APP0 marker detected,
Rem                                0 - if not
Rem    jfif_app0_version  IN/OUT   The JFIF file version
Rem    jfif_app0_units    IN/OUT   units for the X and Y densities
Rem                                0 - no units, X and Y specify
Rem                                    the pixel aspect ratio
Rem                                1 - X and Y are dots per inch
Rem                                2 - X and Y are dots per cm
Rem    jfif_app0_Xdensity IN/OUT   horizontal pixel density
Rem    jfif_app0_Ydensity IN/OUT   vertical pixel density
Rem
Rem    jpeg_comment       IN       pointer to JPEG comments
Rem    jpeg_comment_size  IN/OUT   size of JPEG comments, in bytes
Public Type JPEG_PROPERTIES
  ' Compression/Decompression control.
  iotype                   As IJLIOTYPE                ' default = IJL_SETUP
  roi                      As IJL_RECT                 ' default = 0
  dct_type                 As DCTTYPE                  ' default = IJL_AAN
  fast_processing          As FAST_MCU_PROCESSING_TYPE ' default = IJL_NO_CC_OR_US
  interrupt                As Long                     ' default = FALSE
    
  ' DIB specific I/O data specifiers.
  DIBBytes                 As Long                     ' default = NULL
  DIBWidth                 As Long                     ' default = 0
  DIBHeight                As Long                     ' default = 0
  DIBPadBytes              As Long                     ' default = 0
  DIBChannels              As Long                     ' default = 3
  DIBColor                 As IJL_COLOR                ' default = IJL_BGR
  DIBSubsampling           As IJL_DIBSUBSAMPLING       ' default = IJL_NONE
  DIBLineBytes             As Long                     ' default = 0
    
  ' JPEG specific I/O data specifiers.
  JPGFile                  As Long                     ' default = NULL
  JPGBytes                 As Long                     ' default = NULL
  JPGSizeBytes             As Long                     ' default = 0
  JPGWidth                 As Long                     ' default = 0
  JPGHeight                As Long                     ' default = 0
  JPGChannels              As Long                     ' default = 3
  JPGColor                 As IJL_COLOR                ' default = IJL_YCBCR
  JPGSubsampling           As IJL_JPGSUBSAMPLING       ' default = IJL_411
  JPGThumbWidth            As Long                     ' default = 0
  JPGThumbHeight           As Long                     ' default = 0
    
  ' JPEG conversion properties.
  cconversion_reqd         As Long                     ' default = TRUE
  upsampling_reqd          As Long                     ' default = TRUE
  jquality                 As Long                     ' default = 75
  jinterleaveType          As Long                     ' default = 0
  numxMCUs                 As Long                     ' default = 0
  numyMCUs                 As Long                     ' default = 0
    
  ' Tables.
  nqtables                 As Long
  maxquantindex            As Long
  nhuffActables            As Long
  nhuffDctables            As Long
  maxhuffindex             As Long
    
  jFmtQuant(0 To 3)        As QUANT_TABLE
  jFmtAcHuffman(0 To 3)    As HUFFMAN_TABLE
  jFmtDcHuffman(0 To 3)    As HUFFMAN_TABLE
    
  jEndFmtQuant(0 To 3)     As Long
  jEncFmtAcHuffman(0 To 3) As Long
  jEndFmtDcHuffman(0 To 3) As Long
    
  ' Allow user-defined tables.
  use_external_qtables     As Long
  use_external_htables     As Long
    
  rawquanttables(0 To 3)   As JPEGQuantTable
  rawhufftables(0 To 7)    As JPEGHuffTable
  HuffIdentifierAC(0 To 3) As Byte
  HuffIdentifierDC(0 To 3) As Byte
    
  ' Frame specific members.
  jframe                   As ijl_FRAME
  needframe                As Long
    
  ' SCAN persistent members.
  jscan                    As Long
    
  ' IJL use 8 byte pack structures
  pad0                     As Byte
  pad1                     As Byte
  pad2                     As Byte
  pad3                     As Byte
  
  ' State members.
  state_field              As STATE
  SawAdobeMarker           As Long
  AdobeXform               As Long
    
  ' ROI decoder members.
  rowoffsets               As Long
    
  ' Intermediate buffers.
  MCUBuf                   As Long
  tMCUBuf(0 To 1439)       As Byte
    
  ' Processor detected.
  processortype            As PROCESSOR_TYPE
    
  ' Pointer to array of pointers to buffers for raw DCT coefs.
  raw_coefs                As Long
    
  ' Progressive mode members.
  progressive_found        As Long
  coef_buffer              As Long

  ' Upsampling mode members.
  upsampling_type          As upsampling_type
  sampling_state_ptr       As Long

  ' Adobe APP14 segment variables
  AdobeVersion             As Integer
  AdobeFlags0              As Integer
  AdobeFlags1              As Integer

  ' JFIF APP0 segment variables
  jfif_app0_detected       As Long
  jfif_app0_version        As Integer
  jfif_app0_units          As Byte
  jfif_app0_Xdensity       As Integer
  jfif_app0_Ydensity       As Integer

  ' comments related fields
  jpeg_comment             As Long
  jpeg_comment_size        As Integer

End Type


Rem  Purpose:     This is the primary data structure between the IJL and
Rem               the external user.  It stores JPEG state information
Rem               and controls the IJL.  It is user-modifiable.
Rem  Context:     Used by all low-level IJL routines to store
Rem               pseudo-global information.
Rem  Fields:
Rem    UseJPEGPROPERTIES   Set this flag != 0 if you wish to override
Rem                        the JPEG_CORE_PROPERTIES "IN" parameters with
Rem                        the JPEG_PROPERTIES parameters.
Rem    DIBBytes            IN:     Pointer to buffer of uncompressed data.
Rem    DIBWidth            IN:     Width of uncompressed data.
Rem    DIBHeight           IN:     Height of uncompressed data.
Rem    DIBPadBytes         IN:     Padding (in bytes) at end of each row in the uncompressed data.
Rem    DIBChannels         IN:     Number of components in the uncompressed data.
Rem    DIBColor            IN:     Color space of uncompressed data.
Rem    DIBSubsampling      IN:     Required to be IJL_NONE or IJL_422.
Rem
Rem    JPGFile             IN:     Pointer to file based JPEG.
Rem    JPGBytes            IN:     Pointer to buffer based JPEG.
Rem    JPGSizeBytes        IN:     Max buffer size. Used with JPGBytes.
Rem                        OUT:    Number of compressed bytes written.
Rem    JPGWidth            IN:     Width of JPEG image.
Rem                        OUT:    After reading (except READHEADER).
Rem    JPGHeight           IN:     Height of JPEG image.
Rem                        OUT:    After reading (except READHEADER).
Rem    JPGChannels         IN:     Number of components in JPEG image.
Rem                        OUT:    After reading (except READHEADER).
Rem    JPGColor            IN:     Color space of JPEG image.
Rem    JPGSubsampling      IN:     Subsampling of JPEG image.
Rem                        OUT:    After reading (except READHEADER).
Rem    JPGThumbWidth       OUT:    JFIF embedded thumbnail width [0-255].
Rem    JPGThumbHeight      OUT:    JFIF embedded thumbnail height [0-255].
Rem
Rem    cconversion_reqd    OUT:    If color conversion done on decode, TRUE.
Rem    upsampling_reqd     OUT:    If upsampling done on decode, TRUE.
Rem    jquality            IN:     [0-100] where highest quality is 100.
Rem
Rem    jprops              "Low-Level" IJL data structure.
Rem
Rem//////////////////////////////////////////////////////////////////////////

Public Type JPEG_CORE_PROPERTIES
  UseJPEGPROPERTIES As Long               ' default = 0
    
  ' DIB specific I/O data specifiers.
  DIBBytes          As Long               ' default = NULL
  DIBWidth          As Long               ' default = 0
  DIBHeight         As Long               ' default = 0
  DIBPadBytes       As Long               ' default = 0
  DIBChannels       As Long               ' default = 3
  DIBColor          As IJL_COLOR          ' default = IJL_BGR
  DIBSubsampling    As IJL_DIBSUBSAMPLING ' default = IJL_NONE
    
  ' JPEG specific I/O data specifiers.
  JPGFile           As Long               ' default = NULL
  JPGBytes          As Long               ' default = NULL
  JPGSizeBytes      As Long               ' default = 0
  JPGWidth          As Long               ' default = 0
  JPGHeight         As Long               ' default = 0
  JPGChannels       As Long               ' default = 3
  JPGColor          As IJL_COLOR          ' default = IJL_YCBCR
  JPGSubsampling    As IJL_JPGSUBSAMPLING ' default = IJL_411
  JPGThumbWidth     As Long               ' default = 0
  JPGThumbHeight    As Long               ' default = 0
    
  ' JPEG conversion properties.
  cconversion_reqd  As Long               ' default = TRUE
  upsampling_reqd   As Long               ' default = TRUE
  jquality          As Long               ' default = 75
  
  ' IJL use 8 byte pack structures
  pad0              As Byte
  pad1              As Byte
  pad2              As Byte
  pad3              As Byte
  
  ' Low-level properties.
  jprops            As JPEG_PROPERTIES
End Type


Rem  Purpose:     Listing of possible "error" codes returned by the IJL.
Public Enum IJLERR
  ' The following "error" values indicate an "OK" condition.
  IJL_OK = 0
  IJL_INTERRUPT_OK = 1
  IJL_ROI_OK = 2
    
  ' The following "error" values indicate an error has occurred.
  IJL_EXCEPTION_DETECTED = -1
  IJL_INVLAID_ENCODER = -2
  IJL_UNSUPPORTED_SUBSAMPLING = -3
  IJL_UNSUPPORTED_BYTES_PER_PIXEL = -4
  IJL_MEMORY_ERROR = -5
  IJL_BAD_HUFFMAN_TABLE = -6
  IJL_BAD_QUANT_TABLE = -7
  IJL_INVALID_JPEG_PROPERTIES = -8
  IJL_ERR_FILECLOSE = -9
  IJL_INVALID_FILENAME = -10
  IJL_ERROR_EOF = -11
  IJL_PROG_NOT_SUPPORTED = -12
  IJL_ERR_NOT_JPEG = -13
  IJL_ERR_COMP = -14
  IJL_ERR_SOF = -15
  IJL_ERR_DNL = -16
  IJL_ERR_NO_HUF = -17
  IJL_ERR_NO_QUAN = -18
  IJL_ERR_NO_FRAME = -19
  IJL_ERR_MULT_FRAME = -20
  IJL_ERR_DATA = -21
  IJL_ERR_NO_IMAGE = -22
  IJL_FILE_ERROR = -23
  IJL_INTERNAL_ERROR = -24
  IJL_BAD_RST_MARKER = -25
  IJL_THUMBNAIL_DIB_TOO_SMALL = -26
  IJL_THUMBNAIL_DIB_WRONG_COLOR = -27
  IJL_BUFFER_TOO_SMALL = -28
  IJL_UNSUPPORTED_FRAME = -29
  IJL_ERR_COM_BUFFER = -30
  IJL_RESERVED = -99
End Enum


Public Declare Function ijlInit Lib "ijl15.dll" (ByRef jcprops As JPEG_CORE_PROPERTIES) As IJLERR
Public Declare Function ijlFree Lib "ijl15.dll" (ByRef jcprops As JPEG_CORE_PROPERTIES) As IJLERR
Public Declare Function ijlRead Lib "ijl15.dll" (ByRef jcprops As JPEG_CORE_PROPERTIES, ByVal iotype As IJLIOTYPE) As IJLERR
Public Declare Function ijlWrite Lib "ijl15.dll" (ByRef jcprops As JPEG_CORE_PROPERTIES, ByVal iotype As IJLIOTYPE) As IJLERR
Public Declare Function ijlGetLibVersion Lib "ijl15.dll" () As Long 'pointer to IJLibVersion...
Public Declare Function ijlErrorStr Lib "ijl15.dll" (ByVal code As IJLERR) As Long 'pointer to C-style string


Public Function IJL_DIB_PAD_BYTES(ByVal width As Long, ByVal nChannels As Long) As Long
'Calculate number of bytes to pad DIB line
Dim IJL_DIB_ALIGN As Long
Dim IJL_DIB_UWIDTH As Long
Dim IJL_DIB_AWIDTH As Long

IJL_DIB_ALIGN = 3
IJL_DIB_UWIDTH = width * nChannels
IJL_DIB_AWIDTH = (IJL_DIB_UWIDTH + IJL_DIB_ALIGN) And (Not (IJL_DIB_ALIGN))
IJL_DIB_PAD_BYTES = IJL_DIB_AWIDTH - IJL_DIB_UWIDTH
End Function



Public Function SaveJPGToFile(ByRef cDib As cDIBSection, _
                              ByVal sFile As String) As Boolean
'--------------------------------------------------------------
'Save picture as JPEG image to the file
'this function was not originally part of this module
'--------------------------------------------------------------
Dim jerr As IJLERR
Dim jcprops As JPEG_CORE_PROPERTIES
Dim aFile As String
'On Error Resume Next

jerr = ijlInit(jcprops)
If jerr = IJL_OK Then
   ' Set up the DIB information:
   jcprops.DIBWidth = cDib.dib_width
   jcprops.DIBHeight = -cDib.dib_height
   jcprops.DIBChannels = cDib.dib_channels
   If jcprops.DIBChannels = 3 Then
     jcprops.DIBColor = IJL_BGR
     jcprops.JPGColor = IJL_YCBCR
     jcprops.JPGChannels = 3
     jcprops.JPGSubsampling = IJL_411
   Else
     jcprops.DIBColor = IJL_RGBA_FPX
     jcprops.JPGColor = IJL_YCBCRA_FPX
     jcprops.JPGChannels = 4
     jcprops.JPGSubsampling = IJL_4114
   End If
   ' DIBBytes (pointer to uncompressed RGB data):
   jcprops.DIBBytes = cDib.DIBSectionBitsPtr
   jcprops.DIBPadBytes = IJL_DIB_PAD_BYTES(jcprops.DIBWidth, jcprops.DIBChannels)
   ' Set up the JPEG information:
   aFile = StrConv(sFile, vbFromUnicode)
   jcprops.JPGFile = StrPtr(aFile)
   jcprops.JPGWidth = cDib.dib_width
   jcprops.JPGHeight = cDib.dib_height
   jcprops.jquality = 75
   jerr = ijlWrite(jcprops, IJL_JFILE_WRITEWHOLEIMAGE)
   If jerr = IJL_OK Then
     SaveJPGToFile = True
   Else
    'write to unexpected errors log
    '"Failed to save to JPG"
   End If
   Call ijlFree(jcprops)
Else
   'write to unexpected errors log
   '"Failed to initialise the IJL library"
End If
End Function



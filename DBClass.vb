'-------------------------------------------------------------------------------
'GrFinger Sample
'(c) 2005 Griaule Tecnologia Ltda.
'http://www.griaule.com
'-------------------------------------------------------------------------------
'
'This sample is provided with "GrFinger Fingerprint Recognition Library" and
'can't run without it. It's provided just as an example of using GrFinger
'Fingerprint Recognition Library and should not be used as basis for any
'commercial product.
'
'Griaule Tecnologia makes no representations concerning either the merchantability
'of this software or the suitability of this sample for any particular purpose.
'
'THIS SAMPLE IS PROVIDED BY THE AUTHOR "AS IS" AND ANY EXPRESS OR
'IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES
'OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED.
'IN NO EVENT SHALL GRIAULE BE LIABLE FOR ANY DIRECT, INDIRECT,
'INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT
'NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
'DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
'THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
'(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
'THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'
'You can download the free version of GrFinger directly from Griaule website.
'
'These notices must be retained in any copies of any part of this
'documentation and/or sample.
'
'-------------------------------------------------------------------------------

' -----------------------------------------------------------------------------------
' Database routines
' -----------------------------------------------------------------------------------

Imports System.Data.OleDb
Imports System.Runtime.InteropServices

' Template data
Public Class TTemplate
    ' Template itself
    Public tpt(GrFingerXLib.GRConstants.GR_MAX_SIZE_TEMPLATE) As Byte
    ' Template size
    Public Size As Long
End Class

' Template list
Public Structure TTemplates
    ' ID
    Public ID As Integer
    ' Template itself
    Public template As TTemplate
End Structure


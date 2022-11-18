
standard exe can place a form in ROT for external use with GetObject (thanks wqweto)
normally public class variables on this form can not be accedded but controls can
also classes from standard vb exe can not be accessed directly 

if we patch the class Object.ObjectType field to add bit 0x800 then this changes
to avoid complex parsing, we get the reference to the Object table from a live class instance
then patch the field.

we then need to reinstantate the class to have the new object type kick in. we could also patch
the ObjectType on disk before starting the exe but its cumbersome ever compile unless part of post build
(not really that hard but much more code than this)

start main
start child

child can not access myClass either as a form.myClass or directly

in main click patch class

now child can access form.MyClass

in main click register direct (would fail before patched)

now child can access myClass directly from GetObject 

This should be a handy way to allow remote scripting of your apps with minimal fuss and without having to compile
as an activeX exe which I dislike. long term stability and possible weird nuances not yet tested 





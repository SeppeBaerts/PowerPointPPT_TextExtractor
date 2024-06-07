using OpenMcdf;
using PowerPointTextExtractor.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
[assembly: InternalsVisibleTo("PPTExtractor")]
namespace Extractor_Engine_Service.Extractors
{
    public class PptExtractor
    {
        //These are default in a PPT file, but should not be included.
        private string[] ExcludingStrings = [];
        public string[] DefaultExcludingStrings { get; }= [
            "*", " ",
            "Click to edit Master text styles\r" +
                "Second level\rThird level\rFourth level\rFifth level",
                "Click to edit Master title style",
        ];



        StringBuilder Content = new();

        public string Extract(string pathToFile, string[]? toExcludeStrings = null)
        {
            Stream s = new FileStream(pathToFile, FileMode.Open);
            return Extract(s, toExcludeStrings);
        }
        /// <summary>
        /// Will extract the text from the PPT document. 
        /// </summary>
        /// <param name="toExtractStream"></param>
        /// <param name="toExcludeStrings">If left empty, it will use the defaults. To get all the text, use an empty array.</param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        /// <exception cref="EncryptionException"></exception>
        public string Extract(Stream toExtractStream, string[]? toExcludeStrings = null)
        { 
            if (toExtractStream.Length == 0)
                return string.Empty;
                ExcludingStrings = toExcludeStrings?? DefaultExcludingStrings;
            using (var compoundFile = new CompoundFile(toExtractStream))
            {
                if (!compoundFile.RootStorage.TryGetStream("PowerPoint Document", out var pptStream))
                    throw new Exception("Could not find the PowerPoint Document stream inside the PPT file");

                if (!compoundFile.RootStorage.TryGetStream("Current User", out var userStream))
                    throw new Exception("Could not find the Current User stream inside the PPT file");

                CurrentUserAtom currentUser = new();
                using (var memoryStream = new MemoryStream(userStream.GetData()))
                using (var binaryReader = new BinaryReader(memoryStream))
                {
                    currentUser.Read(binaryReader);
                    if (currentUser.IsEncrypted)
                        throw new EncryptionException("The PowerPoint file is encrypted, and cannot be read by this program");
                }

                using (var memoryStream = new MemoryStream(pptStream.GetData()))
                using (var binaryReader = new BinaryReader(memoryStream))
                {
                    var userEditAtoms = GetAllUserEditAtoms(binaryReader, currentUser.OffsetToCurrentEdit);
                    List<Record> recs = GetValidRecords(binaryReader, userEditAtoms.SelectMany(u => u.PersistDirectory.PersistDirectoryEntries.SelectMany(p => p.PersistOffsets)).ToList());
                }
            }
            var text = Content.ToString();
            return string.IsNullOrWhiteSpace(text) ? string.Empty : text;
        }
        /// <summary>
        /// Gets all the user edit atoms from the powerpoint document stream and resets the position to 0
        /// </summary>
        /// <param name="reader"></param>
        /// <param name="offset"></param>
        /// <returns>A list of UserEditAtoms</returns>
        internal List<UserEditAtom> GetAllUserEditAtoms(BinaryReader reader, uint offset)
        {
            List<UserEditAtom> userEdits = [];
            UserEditAtom userEdit;
            reader.BaseStream.Position = offset;
            do
            {
                userEdit = new(reader);
                userEdits.Add(userEdit);
                offset = userEdit.OffsetLastEdit;
                reader.BaseStream.Position = offset;
            } while (offset != 0);

            reader.BaseStream.Position = 0;
            return userEdits;
        }
        internal Record ReadRecord(BinaryReader reader)
        {
            Record record = Record.CreateRecord(reader);
            if (record is Container container)
            {
                //skip these containers because they do not have the information we need, but they are using CTStrings, so this will make the string content dirty
                if (container.Header.Type == RecordType.RT_ProgBinaryTag || container.Header.Type == RecordType.RT_MainMaster)
                {
                    reader.BaseStream.Position += container.Header.Size;
                    return container;
                }

                container.Children = [];
                long maxOffset = reader.BaseStream.Position + container.Header.Size;
                while (reader.BaseStream.Position < maxOffset)
                    container.Children.Add(ReadRecord(reader));

                return container;
            }
            else if (record is Atom atom)
            {
                atom = atom.Header.Type switch
                {
                    RecordType.RT_CString => new RTCString(reader, atom.Header),
                    RecordType.RT_TextCharsAtom => new TextAtom(reader, atom.Header),
                    RecordType.RT_TextBytesAtom => new TextAtom(reader, atom.Header),
                    _ => atom
                };
                if (atom is TextAtom || atom is RTCString) //if it's a text atom, we want to add the content to the string content
                {
                    string atomContent = atom.ToString();
                    if (!ExcludingStrings.Contains(atomContent))
                        Content.Append($"{atomContent}{(atomContent.EndsWith(' ')? "" : " ")}"); //If it's the end of a dia, we want to add a space. if there is already a space, we do not want to add a space.
                    return atom;
                }
                atom.Data = reader.ReadBytes((int)atom.Header.Size); //else we want to read the data from the atom
                return atom;
            }
            return record;
        }
        internal List<Record> GetValidRecords(BinaryReader reader, List<PersistOffsetEntry> entries)
        {
            List<uint> identifiers = [];
            List<Record> records = [];
            foreach (var entry in entries)
            {
                if (identifiers.Contains(entry.PersistId))
                    continue;

                identifiers.Add(entry.PersistId);
                reader.BaseStream.Position = entry.RgPersistOffset;
                records.Add(ReadRecord(reader));
            }
            return records;

        }
    }
    public class EncryptionException(string message) : Exception(message)
    {

    }
}

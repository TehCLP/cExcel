using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace cExcel
{
    public class cert1
    {
    //    public void gen()
    //    {
    //        string subjectName = "test";

    //        byte[] actual = generator(subjectName);
    //        var cert = new System.Security.Cryptography.X509Certificates.X509Certificate2(actual, "password");
    //    }
    //    public byte[] generator(string subjectName)
    //    {
    //        var kpgen = new RsaKeyPairGenerator();

    //        kpgen.Init(new KeyGenerationParameters(new SecureRandom(new CryptoApiRandomGenerator()), 1024));

    //        var kp = kpgen.GenerateKeyPair();

    //        var gen = new X509V3CertificateGenerator();

    //        var certName = new X509Name("CN=" + subjectName);
    //        var serialNo = BigInteger.ProbablePrime(120, new Random());

    //        gen.SetSerialNumber(serialNo);
    //        gen.SetSubjectDN(certName);
    //        gen.SetIssuerDN(certName);
    //        gen.SetNotAfter(DateTime.Now.AddYears(100));
    //        gen.SetNotBefore(DateTime.Now.Subtract(new TimeSpan(7, 0, 0, 0)));
    //        gen.SetSignatureAlgorithm("MD5WithRSA");
    //        gen.SetPublicKey(kp.Public);

    //        gen.AddExtension(
    //            X509Extensions.AuthorityKeyIdentifier.Id,
    //            false,
    //            new AuthorityKeyIdentifier(
    //                SubjectPublicKeyInfoFactory.CreateSubjectPublicKeyInfo(kp.Public),
    //                new GeneralNames(new GeneralName(certName)),
    //                serialNo));

    //        gen.AddExtension(
    //            X509Extensions.ExtendedKeyUsage.Id,
    //            false,
    //            new ExtendedKeyUsage(new ArrayList() { new DerObjectIdentifier("1.3.6.1.5.5.7.3.1") }));

    //        var newCert = gen.Generate(kp.Private);

    //        return DotNetUtilities.ToX509Certificate(newCert).Export(System.Security.Cryptography.X509Certificates.X509ContentType.Pkcs12, "password");
    //    }
    //}
}
